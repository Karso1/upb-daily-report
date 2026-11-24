import argparse
from calendar import monthrange
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP
import fcntl
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from tempfile import TemporaryDirectory
import warnings

from openpyxl import load_workbook


HERE = Path(__file__).resolve().parent
SOURCE = Path(os.environ.get("UPB_SOURCE_DIR", HERE / "data")).expanduser()
NODE_RUNTIME = Path(os.environ.get("UPB_NODE_RUNTIME", Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node")).expanduser()
NODE = NODE_RUNTIME / "bin/node"
REPORT = HERE / "UPB每日数据汇总.xlsx"
CHANNELS = ("Reap", "Passto", "StraitsX")
RECHARGE_CHANNELS = (*CHANNELS, "其他渠道")
CARD_TYPES = ("虚拟卡", "实体卡")
FX = Decimal("7.78")
DATASETS = {
    "cards": ("开卡", "开卡"),
    "recharge": ("充值数据", "充值"),
    **{channel: ("消费数据", channel) for channel in CHANNELS},
}
SUPPLEMENTAL = ("cards", *CHANNELS)


def read_rows(path, fields):
    alternatives = [(field,) if isinstance(field, str) else field for field in fields]
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Workbook contains no default style")
        book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in book.worksheets:
            sheet.reset_dimensions()
            rows = sheet.values
            headers = [str(cell or "").strip() for cell in next(rows, ())]
            selected = [next((name for name in names if name in headers), None) for names in alternatives]
            if None in selected:
                continue
            indices = [headers.index(name) for name in selected]
            for row in rows:
                yield tuple(str(row[i]).strip() if i < len(row) and row[i] is not None else "" for i in indices)
            return
        raise ValueError(f"{path.name} 找不到包含所需字段的工作表")
    finally:
        book.close()


def money(amount):
    return float(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def is_consumer(agent):
    return agent.casefold() == "upay"


def keep_unique(items, key, value, path, kind):
    if key in items and items[key] != value:
        raise ValueError(f"{path.name} 有金额或归属冲突的重复{kind}记录")
    items[key] = value


def recharge(path, days, monthly):
    orders = {}
    covered = set()
    fields = ("订单号", "状态", "卡片渠道", ("所属代理商", "历史代理商"), "实付币种", "实付金额", "创建时间")
    for order, status, channel, agent, currency, value, created in read_rows(path, fields):
        day = created[:10] if monthly else days[0]
        if day not in days:
            continue
        covered.add(day)
        if status != "充值/抵押成功":
            continue
        if currency != "USDT" or not channel or not order:
            raise ValueError(f"{path.name} 有成功充值的币种、渠道或订单号待核对")
        if channel not in CHANNELS:
            channel = "其他渠道"
        item = (channel, is_consumer(agent), Decimal(value))
        keep_unique(orders, (day, order), item, path, "充值")
    totals = defaultdict(Decimal)
    for (day, _), (channel, consumer, amount) in orders.items():
        totals[day, channel, consumer] += amount
    return {
        day: [money(totals[day, channel, group]) for group in (True, False) for channel in RECHARGE_CHANNELS]
        for day in days if not monthly or day in covered
    }


def cards(path, days, monthly):
    groups = [(channel, card_type) for channel in ("Reap", "StraitsX") for card_type in CARD_TYPES] + [("其他渠道", "")]
    counts = defaultdict(set)
    covered = set()
    fields = ("订单号", "卡ID", "渠道名称", "实/虚卡", "卡状态", "开卡成功时间", "实付开卡费", "申请时间")
    for order, card_id, channel, card_type, status, completed, fee, applied in read_rows(path, fields):
        if applied[:10] in days:
            covered.add(applied[:10])
        day = completed[:10] or (applied[:10] if monthly else "")
        if day not in days:
            continue
        covered.add(day)
        if status not in ("开卡成功", "制卡中", "待激活") or not fee or Decimal(fee) <= 0:
            continue
        group = (channel, card_type) if channel in ("Reap", "StraitsX") else ("其他渠道", "")
        if not channel or group not in groups or not (order or card_id):
            raise ValueError(f"{path.name} 有无法归类的付费开卡记录")
        counts[day, group].add(card_id or order)
    return {day: [len(counts[day, group]) for group in groups] for day in days if not monthly or day in covered}


def transactions(path, channel, days, mode):
    currency = "HKD" if channel == "Passto" else "USD"
    agent = {
        "Passto": ("历史所属代理商", "历史代理商名称", "所属代理商"),
        "Reap": ("历史代理商名称", "历史所属代理商", "所属代理商"),
        "StraitsX": ("历史代理商名称", "历史所属代理商", "所属代理商"),
    }
    amount_fields = ("交易原始金额", "交易金额") if channel == "StraitsX" else "交易金额"
    currency_fields = ("交易原始币种", "交易币种") if channel == "StraitsX" else ("交易币种", "交易币种名称")
    fields = (
        ("ID", "父ID(列表中ID列取该字段)"), ("订单类型", "订单类型名"),
        ("收支类型", "收支类型名", "收支类型文本"), currency_fields, amount_fields, agent[channel],
        "交易时间", ("更新时间", "更新时间(交易)", "平台入账时间"),
    )
    orders = {}
    covered = set()
    errors = {}
    rate = FX if currency == "HKD" else Decimal(1)
    for row_id, kind, flow, row_currency, value, owner, occurred, updated in read_rows(path, fields):
        day = occurred[:10] if mode == "monthly" else days[0]
        if day not in days:
            continue
        covered.add(day)
        if flow != "支出" or (mode != "original" and kind != "主订单"):
            continue
        amount = Decimal(value)
        if channel == "Passto" and mode == "original" and row_currency != "HKD":
            # The manually prepared files put the converted amount in this column.
            if abs(Decimal(row_currency) - amount / FX) < Decimal("0.000001"):
                row_currency = "HKD"
        if row_currency != currency or not row_id:
            errors[day] = f"{path.name} 有支出记录的币种或 ID 待核对（{row_currency or '币种为空'}）"
            continue
        trade = (day, row_id)
        item = (amount, is_consumer(owner))
        prior = orders.get(trade)
        if prior and prior[0] != item:
            previous_update = prior[1]
            if mode != "monthly" or not updated or not previous_update or updated == previous_update:
                errors[day] = f"{path.name} 同一交易 ID 有金额或归属冲突"
                continue
        if not prior or updated >= prior[1]:
            orders[trade] = (item, updated)
    totals = defaultdict(Decimal)
    for (day, _), ((amount, consumer), _) in orders.items():
        totals[day, consumer] += amount
    return {
        day: [money((totals[day, True] + totals[day, False]) / rate), money(totals[day, True] / rate)]
        for day in days if day not in errors and (mode != "monthly" or day in covered)
    }, errors


def aggregate(task):
    key, path, mode, days = task
    errors = {}
    try:
        initial = path.stat()
        if key == "cards":
            values = cards(path, days, monthly=mode == "monthly")
        elif key == "recharge":
            values = recharge(path, days, monthly=mode == "monthly")
        else:
            values, errors = transactions(path, key, days, mode)
        final = path.stat()
        if (initial.st_size, initial.st_mtime_ns) != (final.st_size, final.st_mtime_ns):
            raise ValueError(f"{path.name} 在读取时发生变化，请重新运行")
        return key, path, days, values, errors, None
    except (ArithmeticError, OSError, ValueError) as error:
        return key, path, days, {}, {}, f"{path.name} · {error}"


def load_report(path):
    if not path.exists():
        return {}
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book["UPB每日汇总"]
        old_layout = sheet["A34"].value == "交易汇总"
        current_layout = sheet["A11"].value == "其他渠道开卡"
        last_row = 53 if current_layout else 35 if old_layout else 48
        columns = list(zip(*sheet.iter_rows(min_row=2, max_row=last_row, values_only=True)))
        if not columns or columns[0][1] != "平台注册人数":
            raise ValueError("现有 Excel 的平台注册人数行已移动")
        sources = book["数据来源"]
        records = {}
        source_rows = sources.iter_rows(min_row=3, max_row=len(columns) + 1, max_col=8, values_only=True)
        for column, source_row in zip(columns[1:], source_rows, strict=True):
            day, *paths, flags, issues = source_row
            if not isinstance(column[0], datetime) or column[0].date().isoformat() != day or day in records:
                raise ValueError("现有 Excel 的日期与来源记录不一致")
            registration = column[1]
            card_values = list(column[5:10] if current_layout else column[5:9])
            recharge_values = list(column[16:24] if current_layout else column[14:20])
            tx_start = 34 if current_layout else 29
            if not current_layout:
                card_values.append(0 if all(value is not None for value in card_values) else None)
                missing_value = 0 if all(value is not None for value in recharge_values) else None
                recharge_values.insert(3, missing_value)
                recharge_values.append(missing_value)
            records[day] = {
                "day": day, "registration": registration, "card_status": column[2], "tx_status": column[3],
                "cards": card_values, "recharge": recharge_values,
                "transactions": [[column[i], None if old_layout else column[i + 3]] for i in range(tx_start, tx_start + 3)],
                "user_funds": None if old_layout else column[-2],
                "wealth": None if old_layout else column[-1],
                "paths": dict(zip(DATASETS, paths)), "supplemental": flags.split(",") if flags else [],
                "issues": issues or "",
            }
        return records
    finally:
        book.close()


def archive_months(root, archive_year):
    archives = {}

    def register(key, month, path):
        existing = archives.get((key, month))
        if existing is None or (path.stat().st_mtime_ns, path.name) > (existing.stat().st_mtime_ns, existing.name):
            archives[key, month] = path

    for key, folder in [("cards", "开卡"), ("recharge", "充值数据")]:
        for path in (root / folder).rglob("*.xlsx"):
            if path.name.startswith("~") or "每日自动下载" in path.parts:
                continue
            match = re.search(r"(20[0-9]{2})[-年](0?[1-9]|1[0-2])(?:月|[^0-9]|$)", path.stem)
            if match:
                year, month = map(int, match.groups())
                register(key, f"{year:04}-{month:02}", path)

    for directory in (root / "消费数据/UPB").glob("Q*/*月*"):
        month_match = re.fullmatch(r"(\d{1,2})月(?:份)?", directory.name)
        if not month_match:
            continue
        month = int(month_match[1])
        for path in directory.glob("*.xlsx"):
            key = next((channel for channel in CHANNELS if path.name.lower().startswith(channel.lower())), None)
            if not key and path.name.casefold().startswith("straitx"):
                key = "StraitsX"
            if not key:
                continue
            year_match = re.search(r"20[0-9]{2}", path.name)
            year = int(year_match[0]) if year_match else archive_year
            register(key, f"{year:04}-{month:02}", path)
    return archives


def daily_files(root):
    result = {}
    for key, (folder, prefix) in DATASETS.items():
        for path in (root / folder / "每日自动下载").glob(f"*/{prefix}_*.xlsx"):
            iso = path.stem.removeprefix(prefix + "_")
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", iso) and path.parent.name == iso[:7]:
                result[key, iso] = path
    return result


def legacy_files(root):
    selectors = {
        "cards": ("UPB开卡数据", "卡片申请历史"),
        "recharge": ("UPB充值数据", "充值订单"),
        "Reap": ("UPB消费数据/reap", "Reap渠道明细账务"),
        "Passto": ("UPB消费数据/passto", "Passto渠道明细账务"),
        "StraitsX": ("UPB消费数据/straitsx", "StraitsX渠道账务明细"),
    }
    result = {}
    for key, (folder, prefix) in selectors.items():
        for path in (root / folder).glob(f"{prefix}*.xlsx"):
            day = path.stem.removeprefix(prefix)
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", day):
                result[key, day] = path
    return result


def archive_dates(archive):
    days = set()
    for month in {month for _, month in archive}:
        year, number = map(int, month.split("-"))
        days.update(date(year, number, day).isoformat() for day in range(1, monthrange(year, number)[1] + 1))
    return days


def select_sources(day, previous, daily, archive, overrides):
    sources = {}
    for key in DATASETS:
        previous_path = previous.get("paths", {}).get(key)
        original = overrides.get(key)
        if not original and key in previous.get("supplemental", []) and previous_path:
            candidate = Path(previous_path)
            if candidate.is_file():
                original = candidate
        if original:
            sources[key] = (original, "original")
        elif path := daily.get((key, day)):
            sources[key] = (path, "daily")
        elif path := archive.get((key, day[:7])):
            sources[key] = (path, "monthly")
    return sources


def collect(selected):
    requests = defaultdict(list)
    for day, sources in selected.items():
        for key, (path, mode) in sources.items():
            requests[key, path, mode].append(day)
    tasks = [(*source, days) for source, days in requests.items()]
    tasks.sort(key=lambda task: task[1].stat().st_size, reverse=True)
    amounts = defaultdict(dict)
    issues = defaultdict(dict)
    if not tasks:
        return amounts, issues
    print(f"读取 {len(tasks)} 份 UPB 源表，检查 {len(selected)} 个日期。", flush=True)
    with ProcessPoolExecutor(max_workers=min(6, len(tasks))) as pool:
        for index, (key, path, days, values, day_errors, error) in enumerate(pool.map(aggregate, tasks), 1):
            if error:
                print(error, flush=True)
            for day, message in day_errors.items():
                print(f"{day} · {message}", flush=True)
            for day in days:
                if day in values:
                    amounts[day][key] = values[day]
                else:
                    issues[day][key] = error or day_errors.get(day) or f"{path.name} 当天无明细"
            if len(tasks) > 10 and (index % 10 == 0 or index == len(tasks)):
                print(f"已读取 {index}/{len(tasks)} 份。", flush=True)
    return amounts, issues


def report_row(day, selected, amounts, issues, previous):
    mode = {key: source[1] for key, source in selected.items()}
    paths = {key: str(selected[key][0]) if key in selected else None for key in DATASETS}
    card_status = {
        "original": "完成日导出", "daily": "待核对 · 申请日导出", "monthly": "待核对 · 月度归档",
        "legacy": "待核对 · 旧版日档",
    }.get(mode.get("cards"), "缺少开卡源表")
    if "cards" in issues:
        card_status = "开卡源表待核对"
    if any(key not in mode for key in CHANNELS):
        tx_status = "缺少交易源表"
    elif any(key in issues for key in CHANNELS):
        tx_status = "交易源表待核对"
    elif all(mode[key] == "original" for key in CHANNELS):
        tx_status = "原日报导出"
    elif all(mode[key] == "monthly" for key in CHANNELS):
        tx_status = "待核对 · 月度归档"
    elif "legacy" in (mode[key] for key in CHANNELS):
        tx_status = "待核对 · 旧版日档"
    else:
        tx_status = "待核对 · 自动导出"
    return {
        "day": day, "registration": previous.get("registration"), "user_funds": previous.get("user_funds"),
        "wealth": previous.get("wealth"), "card_status": card_status, "tx_status": tx_status,
        "cards": amounts.get("cards", [None] * 5), "recharge": amounts.get("recharge", [None] * 8),
        "transactions": [amounts.get(key, [None, None]) for key in CHANNELS],
        "paths": paths, "supplemental": [key for key in DATASETS if mode.get(key) == "original"],
        "issues": "；".join(issues.values()),
    }


def main():
    parser = argparse.ArgumentParser(description="更新本地 UPB 横向日报")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--date", type=date.fromisoformat, help="重算指定日期 YYYY-MM-DD")
    mode.add_argument("--history", action="store_true", help="从现有月度归档和每日源表重算历史")
    parser.add_argument("--from", type=date.fromisoformat, dest="start_date", help="历史起始日期 YYYY-MM-DD")
    parser.add_argument("--to", type=date.fromisoformat, dest="end_date", help="历史结束日期 YYYY-MM-DD")
    parser.add_argument("--registration", type=int, help="指定日期的平台注册人数")
    parser.add_argument("--source", type=Path, default=SOURCE)
    parser.add_argument("--legacy-source", type=Path, default=os.environ.get("UPB_LEGACY_SOURCE_DIR"),
                        help="可选的旧版每日数据库")
    parser.add_argument("--archive-year", type=int, default=datetime.now(timezone.utc).year,
                        help="消费月档文件名未注明年份时使用的年份")
    parser.add_argument("--output", type=Path, default=REPORT)
    parser.add_argument("--preview", type=Path)
    for key in SUPPLEMENTAL:
        parser.add_argument(f"--{key.lower()}-export", type=Path, dest=key)
    args = parser.parse_args()
    overrides = {key: getattr(args, key).resolve() for key in SUPPLEMENTAL if getattr(args, key)}
    if (overrides or args.registration is not None) and not args.date:
        parser.error("填写注册人数或补充导出时，请同时指定 --date")
    if args.registration is not None and args.registration < 0:
        parser.error("注册人数必须是非负整数")
    if (args.start_date or args.end_date) and not args.history:
        parser.error("--from 和 --to 用于 --history")
    if args.start_date and args.end_date and args.start_date > args.end_date:
        parser.error("起始日期须早于或等于结束日期")
    if any(not path.is_file() for path in overrides.values()):
        parser.error("找不到指定的补充导出文件")
    if not args.source.is_dir():
        parser.error(f"找不到 UPB 原始数据文件夹 {args.source}")
    if args.legacy_source and not args.legacy_source.expanduser().is_dir():
        parser.error(f"找不到旧版每日数据库 {args.legacy_source}")
    if not 2000 <= args.archive_year <= 2099:
        parser.error("归档年份须在 2000 至 2099 之间")

    args.output = args.output.resolve()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.with_name(f".{args.output.name}.lock").open("w") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        update(args, overrides)


def update(args, overrides):
    if args.output.with_name(f"~${args.output.name}").exists():
        raise ValueError("请先保存并关闭汇总 Excel，再运行更新")
    last_edit = args.output.stat().st_mtime_ns if args.output.exists() else None
    records = load_report(args.output)
    root = args.source.resolve()
    daily = daily_files(root)
    archive = archive_months(root, args.archive_year) if args.history or args.date else {}
    legacy = legacy_files(args.legacy_source.expanduser().resolve()) if args.legacy_source and (args.history or args.date) else {}
    completed = datetime.now(timezone.utc).date() - timedelta(days=1)
    daily_dates = {day for _, day in daily if day <= completed.isoformat()}
    if args.date:
        requested = [args.date.isoformat()]
    elif args.history:
        start = args.start_date.isoformat() if args.start_date else ""
        end = (args.end_date or completed).isoformat()
        requested = sorted(day for day in daily_dates | archive_dates(archive) | {iso for _, iso in legacy} | set(records) if start <= day <= end)
    else:
        latest = sorted(daily_dates | {day for day in records if day <= completed.isoformat()})[-2:]
        requested = sorted(set(latest) | {day for day in daily_dates if records and day > max(records)})

    selected = {day: select_sources(day, records.get(day, {}), daily, archive, overrides) for day in requested}
    amounts, issues = collect(selected)
    fallback = {
        day: {
            key: (legacy[key, day], "legacy") for key in DATASETS
            if (key, day) in legacy and key not in amounts[day]
            and (key not in selected[day] or selected[day][key][1] == "monthly")
        }
        for day in requested
    }
    fallback = {day: files for day, files in fallback.items() if files}
    if fallback:
        print("用旧版每日源表补齐月档缺失日期。", flush=True)
        fallback_amounts, fallback_issues = collect(fallback)
        for day, datasets in fallback_amounts.items():
            for key, value in datasets.items():
                selected[day][key] = fallback[day][key]
                amounts[day][key] = value
                issues[day].pop(key, None)
        for day, datasets in fallback_issues.items():
            for key, message in datasets.items():
                issues[day].setdefault(key, message)
    changed = []
    for day in requested:
        record = report_row(day, selected[day], amounts[day], issues[day], records.get(day, {}))
        if args.registration is not None:
            record["registration"] = args.registration
        if records.get(day) != record:
            records[day] = record
            changed.append(day)

    if not records or (not changed and not args.preview):
        print("当前汇总数据一致。" if records else "此范围内没有 UPB 源表。")
        return
    with TemporaryDirectory(prefix=".upb-", dir=args.output.parent) as temp:
        input_json = Path(temp) / "daily.json"
        input_json.write_text(json.dumps([records[key] for key in sorted(records)], ensure_ascii=False), encoding="utf-8")
        candidate = Path(temp) / args.output.name
        render = [str(NODE), str(HERE / "build.mjs"), str(input_json), str(candidate)]
        if args.preview:
            render.append(str(args.preview.resolve()))
        if not NODE.is_file():
            raise ValueError("找不到 Node.js，请设置 UPB_NODE_RUNTIME 指向包含 bin/node 的运行时目录")
        subprocess.run(render, check=True, env={**os.environ, "UPB_NODE_RUNTIME": str(NODE_RUNTIME.resolve())})
        current_edit = args.output.stat().st_mtime_ns if args.output.exists() else None
        if last_edit != current_edit or args.output.with_name(f"~${args.output.name}").exists():
            raise ValueError("汇总 Excel 在更新期间被打开或修改，请保存并关闭后重试")
        candidate.replace(args.output)
    if changed:
        print(f"已更新 {len(changed)} 天（{changed[0]} 至 {changed[-1]}）。")
    else:
        print("已生成预览。")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, ArithmeticError, subprocess.CalledProcessError) as error:
        sys.exit(f"更新失败。{error}")
