from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

from xlsxwriter import Workbook
from xlsxwriter.utility import xl_rowcol_to_cell


PHOTON = "PhotonPay"
BASE_CHANNELS = ("Reap", "Passto", "StraitsX")
CHANNELS = (*BASE_CHANNELS, PHOTON)
RECHARGE_CHANNELS = (*CHANNELS, "其他渠道")
FX = Decimal("7.78")
NAVY = "#12323B"
TEAL = "#08786C"
PALE = "#E8F5F2"
AMBER = "#FFF0CE"

LABELS = (
    "UP Business 卡业务日汇总", "报表日期（UTC+0）", "平台注册人数", "开卡来源", "交易来源",
    "金融卡", "Reap 虚拟卡", "Reap 实体卡", "StraitsX 虚拟卡", "StraitsX 实体卡", "其他渠道开卡", "开卡汇总",
    "抵押 / 充值（USDT）", *RECHARGE_CHANNELS, *(c + " 2C" for c in RECHARGE_CHANNELS),
    *(c + " 2B" for c in RECHARGE_CHANNELS), *(c + " 充值占比" for c in RECHARGE_CHANNELS),
    "2B 充值", "2C 充值", "2B 充值占比", "2C 充值占比", "充值汇总",
    "交易明细（USD）", *CHANNELS, *(c + " 2C" for c in CHANNELS), *(c + " 2B" for c in CHANNELS),
    "2B 消费", "2C 消费", "2B 消费占比", "2C 消费占比", "消费汇总", "Passto 汇率（HKD/USD）",
    "资金明细（USD）", "用户资金", "理财金额",
)
ROW = {label: index for index, label in enumerate(LABELS, 1)}
TX_START = ROW["交易明细（USD）"] + 1
RECHARGE_2C = 14 + len(RECHARGE_CHANNELS)
RECHARGE_2B = RECHARGE_2C + len(RECHARGE_CHANNELS)
RECHARGE_SHARE = RECHARGE_2B + len(RECHARGE_CHANNELS)
TX_2C = TX_START + len(CHANNELS)
TX_2B = TX_2C + len(CHANNELS)
SOURCE_HEADERS = ("日期", "开卡", "充值", *CHANNELS, "补充导出", "待核对明细", "PhotonPay 预计有源表")
SOURCE_NOTES = (
    "平台注册人数、用户资金、理财金额为黄色手填行，按日期保留已保存的数字。",
    "开卡只计实付费大于 0 的开卡成功、制卡中、待激活。按完成日归集，旧月档缺完成时间时按申请日归集。",
    "充值取充值/抵押成功订单的 USDT 实付金额。充值按所属代理商、交易按渠道导出中的历史代理商分组，UPay 归 2C，其余归 2B。",
    "交易按支出汇总，日报补充导出沿用其原有范围。CRM 日档和归档仅统计主订单。Passto 采用 7.78 HKD/USD。",
    "光子易和 PhotonPay 统一记为 PhotonPay，消费按 USD 交易金额统计。已覆盖期间缺源表时消费汇总留空。",
    "历史计算优先用补充导出、CRM 日档、月档，旧版日档补月档漏缺。月档充值按创建日、交易按交易日归集。",
    "历史开卡和充值中的 OE 等渠道归入其他渠道。月档中同一交易 ID 有多次更新时使用较新的记录。",
    "缺少源表或相关日期的明细时对应金额或张数留空。黄色来源标记表示导出范围待核对。",
)


def calculate_column(record, column):
    """Return input values, Excel formulas, and their initial cached results."""
    values = {
        2: datetime.fromisoformat(record["day"]),
        3: record["registration"],
        4: record["card_status"],
        5: record["tx_status"],
        ROW["Passto 汇率（HKD/USD）"]: float(FX) if record["transactions"][1][0] is not None else None,
        ROW["用户资金"]: record["user_funds"],
        ROW["理财金额"]: record["wealth"],
        **dict(enumerate(record["cards"], 7)),
        **dict(enumerate(record["recharge"], RECHARGE_2C)),
    }
    formulas = {}

    def cell(row):
        return xl_rowcol_to_cell(row - 1, column)

    def arithmetic(row, inputs, expression, operation=sum, refs=None):
        numbers = [values[r] for r in inputs]
        values[row] = (
            float(operation([Decimal(str(number)) for number in numbers]).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
            if all(number is not None for number in numbers) else None
        )
        count_range = refs or ",".join(cell(r) for r in inputs)
        formulas[row] = f'=IF(COUNT({count_range})={len(inputs)},ROUND({expression},2),"")'

    def subtotal(row, first, last):
        refs = f"{cell(first)}:{cell(last)}"
        arithmetic(row, range(first, last + 1), f"SUM({refs})", refs=refs)

    def share(row, amount, total):
        numerator, denominator = values[amount], values[total]
        values[row] = numerator / denominator if numerator is not None and denominator is not None and denominator > 0 else None
        a, t = cell(amount), cell(total)
        formulas[row] = f'=IF(AND(ISNUMBER({a}),ISNUMBER({t}),{t}>0),{a}/{t},"")'

    subtotal(12, 7, 11)
    r2b, r2c, rt = (ROW[k] for k in ("2B 充值", "2C 充值", "充值汇总"))
    for i in range(len(RECHARGE_CHANNELS)):
        a, b = RECHARGE_2C + i, RECHARGE_2B + i
        arithmetic(14 + i, (a, b), f"SUM({cell(a)},{cell(b)})")
    subtotal(r2b, RECHARGE_2B, RECHARGE_SHARE - 1)
    subtotal(r2c, RECHARGE_2C, RECHARGE_2B - 1)
    subtotal(rt, r2b, r2c)
    for i in range(len(RECHARGE_CHANNELS)):
        share(RECHARGE_SHARE + i, 14 + i, rt)
    share(ROW["2B 充值占比"], r2b, rt)
    share(ROW["2C 充值占比"], r2c, rt)

    for i, (total, consumer) in enumerate(record["transactions"]):
        a, b = TX_START + i, TX_2C + i
        values[a], values[b] = total, consumer
        arithmetic(TX_2B + i, (a, b), f"{cell(a)}-{cell(b)}", operation=lambda amounts: amounts[0] - amounts[1])
    t2b, t2c, tt = (ROW[k] for k in ("2B 消费", "2C 消费", "消费汇总"))
    last = (len(CHANNELS) if record["photon_required"] else len(BASE_CHANNELS)) - 1
    subtotal(t2b, TX_2B, TX_2B + last)
    subtotal(t2c, TX_2C, TX_2C + last)
    subtotal(tt, TX_START, TX_START + last)
    share(ROW["2B 消费占比"], t2b, tt)
    share(ROW["2C 消费占比"], t2c, tt)
    return values, formulas


def summary_formats(workbook):
    headers = {1, 6, 13, ROW["交易明细（USD）"], ROW["资金明细（USD）"]}
    totals = {2, 12, *(ROW[label] for label in ("2B 充值", "2C 充值", "充值汇总", "2B 消费", "2C 消费", "消费汇总"))}
    manual = {3, ROW["用户资金"], ROW["理财金额"]}
    integers = {3, *range(7, 13)}
    percentages = {i for i, label in enumerate(LABELS, 1) if label.endswith("占比")}
    formats = {}
    for row in range(1, len(LABELS) + 1):
        style = {"font_name": "Arial", "font_size": 11, "font_color": NAVY, "valign": "vcenter"}
        if row in headers:
            style.update(bg_color=NAVY if row == 1 else TEAL, font_color="#FFFFFF", bold=True)
        elif row in totals:
            style.update(bg_color=PALE, bold=True)
        elif row in manual:
            style["bg_color"] = AMBER
        if row not in headers and row not in (4, 5):
            style["num_format"] = (
                "yyyy-mm-dd" if row == 2 else "#,##0" if row in integers
                else "0.00%" if row in percentages else "#,##0.00"
            )
        formats[row] = workbook.add_format(style)
    return formats


def write_summary(workbook, records):
    sheet = workbook.add_worksheet("UPB每日汇总")
    sheet.hide_gridlines(2)
    sheet.freeze_panes(2, 1)
    sheet.set_column_pixels(0, 0, 276)
    sheet.set_column_pixels(1, len(records), 156)
    formats = summary_formats(workbook)
    title = workbook.add_format({
        "font_name": "Arial", "font_size": 16, "font_color": "#FFFFFF", "bg_color": NAVY,
        "bold": True, "valign": "vcenter",
    })
    status_style = {"font_name": "Arial", "font_size": 10, "font_color": NAVY, "valign": "vcenter"}
    status = workbook.add_format(status_style)
    attention = workbook.add_format({**status_style, "bg_color": AMBER})
    for row, label in enumerate(LABELS, 1):
        sheet.set_row(row - 1, 35 if row == 1 else 24 if row in (6, 13, ROW["交易明细（USD）"], ROW["资金明细（USD）"]) else 21)
        sheet.write_string(row - 1, 0, label, title if row == 1 else formats[row])

    for column, record in enumerate(records, 1):
        values, formulas = calculate_column(record, column)
        for row, cell_format in formats.items():
            value = values.get(row)
            if row in (4, 5):
                verified = "完成日导出" if row == 4 else "原日报导出"
                cell_format = status if value == verified else attention
            if row in formulas:
                sheet.write_formula(row - 1, column, formulas[row], cell_format, "" if value is None else value)
            else:
                sheet.write(row - 1, column, value, cell_format)


def write_sources(workbook, records):
    sheet = workbook.add_worksheet("数据来源")
    sheet.hide_gridlines(2)
    sheet.freeze_panes(2, 0)
    sheet.set_column_pixels(0, 0, 108)
    sheet.set_column_pixels(1, 6, 220)
    sheet.set_column_pixels(7, 7, 172)
    sheet.set_column_pixels(8, 8, 300)
    sheet.set_column_pixels(9, 9, 172)
    body_style = {"font_name": "Arial", "font_size": 10, "font_color": NAVY, "valign": "vcenter"}
    header = workbook.add_format({**body_style, "bold": True, "bg_color": NAVY, "font_color": "#FFFFFF"})
    body = workbook.add_format({**body_style, "text_wrap": True})
    sheet.set_row(0, 25)
    sheet.set_row(1, 25)
    sheet.merge_range("A1:J1", "UPB 日报数据来源", header)
    sheet.write_row(1, 0, SOURCE_HEADERS, header)
    for row, record in enumerate(records, 2):
        paths = record["paths"]
        values = (record["day"], paths["cards"], paths["recharge"], *(paths[c] for c in CHANNELS),
                  ",".join(record["supplemental"]), record["issues"], record["photon_required"])
        sheet.set_row(row, 72)
        sheet.write_row(row, 0, values, body)
    for row, note in enumerate(SOURCE_NOTES, len(records) + 3):
        sheet.set_row(row, 22)
        sheet.merge_range(row, 0, row, 9, note, body)


def write_report(output, records):
    options = {"strings_to_formulas": False, "strings_to_urls": False}
    with Workbook(str(output), options) as workbook:
        write_summary(workbook, records)
        write_sources(workbook, records)
