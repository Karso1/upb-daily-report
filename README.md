# UPB 每日数据汇总 / UPB Daily Report

[简体中文](#简体中文) · [English](#english)

## 简体中文

UPB 日报把已经下载到本地的 CRM Excel 汇总成一个工作簿，每个 UTC 日期一列。它是下载后的处理步骤，整个计算和 Excel 输出均在本机完成，无需 CRM、Lark 或 GitHub API 凭据。

### 数据从哪里来

`--source` 指定 UPB 源数据目录。公开版默认读取项目下的 `data/`，也可以设置 `UPB_SOURCE_DIR`。每日文件按下面的相对路径识别，文件名中的日期是导出批次的报表日。

| 指标 | CRM 导出 | 相对于源数据目录的路径 |
| --- | --- | --- |
| 开卡 | 开卡申请历史 | `开卡/每日自动下载/YYYY-MM/开卡_YYYY-MM-DD.xlsx` |
| 充值 | 充值／抵押订单 | `充值数据/每日自动下载/YYYY-MM/充值_YYYY-MM-DD.xlsx` |
| Reap 交易 | Reap 渠道明细账务 | `消费数据/每日自动下载/YYYY-MM/Reap_YYYY-MM-DD.xlsx` |
| Passto 交易 | Passto 渠道明细账务 | `消费数据/每日自动下载/YYYY-MM/Passto_YYYY-MM-DD.xlsx` |
| StraitsX 交易 | StraitsX 渠道账务明细 | `消费数据/每日自动下载/YYYY-MM/StraitsX_YYYY-MM-DD.xlsx` |

开卡的 CRM 日档按申请时间导出，充值日档按订单创建时间导出。Passto 和 Reap 的 CRM 交易日档按更新时间导出，StraitsX 日档按交易时间导出。交易日档按整个文件归到文件名对应的日期，因此更新时间口径的日档可能含有更早发生的交易。

历史模式还识别开卡目录下带年月的归档、`充值数据/Q*/` 下带年月的充值月档，以及 `消费数据/UPB/Q*/<月份>/` 下以渠道名开头的交易月档。交易月档文件名无年份时，使用 `--archive-year` 指定的年份，默认当年 UTC 年份。一个渠道同月有多份候选月档时，选择最后修改的文件。

来源按日期、按指标选择。优先级为指定的原日报补充导出、CRM 日档、月档。配置 `--legacy-source` 或 `UPB_LEGACY_SOURCE_DIR` 后，月档缺少相应日期或无法计算时，还会尝试旧版每日数据库。旧日档使用 `UPB开卡数据/`、`UPB充值数据/`、`UPB消费数据/{reap,passto,straitsx}/`，文件名为导出名紧跟 `YYYY-MM-DD.xlsx`。每一日实际选用的文件在 Excel 的「数据来源」页可查。

### 使用方法

需要 Python 3.10+、`requirements.txt` 中的 openpyxl，以及能加载 `@oai/artifact-tool` 的 Node.js 运行时。Excel 输出依赖该库，当前实测环境是 Codex 随附的运行时。其他环境需已有对该库的访问权限。设置 `UPB_NODE_RUNTIME` 可指定包含 `bin/node` 和该库 `node_modules` 的目录。

```sh
python3 -m pip install -r requirements.txt

# 日常更新
python3 refresh.py --source "/path/to/upb-exports"

# 重算所有可识别的历史日期
python3 refresh.py --history --source "/path/to/upb-exports" --legacy-source "/path/to/legacy-exports"

# 重算一个区间，含起止日期
python3 refresh.py --history --from 2026-05-01 --to 2026-05-31 --source "/path/to/upb-exports"

# 重算单日，指定该日按完成时间导出的开卡明细
python3 refresh.py --date 2026-05-12 --source "/path/to/upb-exports" --cards-export "/path/to/card-completions.xlsx"
```

默认输出为项目目录下的 `UPB每日数据汇总.xlsx`，可通过 `--output` 覆盖。交易补充文件分别通过 `--reap-export`、`--passto-export` 和 `--straitsx-export` 指定，须与 `--date` 一起使用。补充文件路径会写进来源页，文件仍在时，重算同一天会复用它。

macOS 双击入口是 `双击更新UPB日报.command` 和 `双击重算UPB历史.command`。如果下载后脚本缺少执行权限，运行 `chmod +x *.command`。双击依赖 Finder 启动时能找到配置好的 Python 和默认数据目录，其他安装方式可用上面的终端命令。更新前先保存并关闭 Excel。

### 实现步骤和技术

1. **确定日期。** 日常模式默认重算最近两个已知 UTC 日期，并追加汇总末尾之后已下载的日期。首次运行默认取有源文件的最近两个日期。中间漏掉的旧日期用历史模式或 `--date` 补。历史模式重新读取所选日期的归档和日档，默认截至 UTC 昨天。
2. **读取源表。** `refresh.py` 使用 openpyxl 的只读模式，查找表头符合要求的工作表。Python 按文件并行读取，解析金额时使用 `Decimal`。读完检查源文件大小和修改时间，遇到来源缺失或数据问题时记录待核对信息。
3. **按日期汇总。** Python 按渠道、客群归集并去重，把各日期的汇总、来源和手填值作为临时 JSON 交给 `build.mjs`。JavaScript 通过 Artifact Tool 生成「UPB每日汇总」和「数据来源」，并写入合计、分组占比等 Excel 公式。
4. **保存。** 更新时从现有日报按日期读回「平台注册人数」「用户资金」「理财金额」，包括手填的 0。已存在的日期原位更新。程序对同一个输出加锁，在临时目录生成新工作簿，再检查旧文件是否被打开或修改，最后替换输出。源 Excel 全程只读。

### 指标怎么算

| 指标 | 规则 |
| --- | --- |
| 开卡 | 只计实付开卡费大于 0，且状态为「开卡成功」「制卡中」或「待激活」的记录。优先按开卡成功时间归日，月档该字段为空时用申请日。按日期、渠道和卡种用卡 ID 去重，缺卡 ID 时用订单号。单份日档若按申请时间导出，可能漏掉早前申请、当天才完成的卡。 |
| 充值（USDT） | 只计状态为「充值/抵押成功」的订单，累加 `实付金额`，要求实付币种为 USDT，按日期和订单号去重。月档按创建时间归日，日档按文件名日期归日。通过 `卡片渠道` 分组，其他渠道计入「其他渠道」。 |
| 消费（USD） | 只计「支出」。CRM 日档、月档及旧版日档限「主订单」，逐日指定的原日报补充文件沿用其筛选范围。Reap 读 `交易金额`（USD），StraitsX 优先读 `交易原始金额`（旧版用 `交易金额`，USD），Passto 读 HKD `交易金额` 后除以固定的 `7.78`。按日期和交易 ID 去重，月档同一 ID 有金额或归属变动且更新时间可区分时取较新的记录。月档按交易时间归日。 |
| 2B／2C | 代理商字段去掉前后空白并忽略大小写后，等于 `UPay` 的归 2C，其余归 2B。充值读取所属代理商，交易优先读取渠道的历史代理商字段。交易的 2B 金额为渠道合计减渠道 2C。 |
| 合计和占比 | 金额按分组小计四舍五入至两位小数。占比为分组或渠道金额除以同日对应总额。分母为零或有必要输入缺失时，留空。 |

汇总页标黄的来源状态表示导出范围待核对。缺档、币种异常或无法消解的重复记录会留下待核对说明，受影响的数字留空。工作簿、原始明细和本地配置属于使用环境中的数据，保存在本地。

## English

UPB Daily Report turns locally downloaded CRM spreadsheets into one Excel workbook with one column per UTC reporting date. It runs after the export step and processes and writes Excel locally. No CRM, Lark or GitHub API credentials are needed.

### Input sources

Pass the UPB export root as `--source` or set `UPB_SOURCE_DIR`. The public checkout defaults to `data/` within the project. Daily inputs are detected by these relative paths. The date in the filename represents the export batch's reporting date.

| Metric | CRM export | Path relative to the export root |
| --- | --- | --- |
| Card openings | Card application history | `开卡/每日自动下载/YYYY-MM/开卡_YYYY-MM-DD.xlsx` |
| Top-ups | Successful top-up / collateral orders | `充值数据/每日自动下载/YYYY-MM/充值_YYYY-MM-DD.xlsx` |
| Reap transactions | Reap channel ledger | `消费数据/每日自动下载/YYYY-MM/Reap_YYYY-MM-DD.xlsx` |
| Passto transactions | Passto channel ledger | `消费数据/每日自动下载/YYYY-MM/Passto_YYYY-MM-DD.xlsx` |
| StraitsX transactions | StraitsX channel ledger | `消费数据/每日自动下载/YYYY-MM/StraitsX_YYYY-MM-DD.xlsx` |

The upstream CRM card export is filtered by application time, and the top-up export by creation time. The Passto and Reap daily exports use update time, while StraitsX uses transaction time. The report assigns each daily transaction file to its filename date, so an update-time batch can contain transactions that took place earlier.

History mode also searches for month-dated card files under `开卡/`, month-dated top-ups under `充值数据/Q*/`, and channel-named ledgers under `消费数据/UPB/Q*/<month>/`. For transaction archive filenames without a year, set `--archive-year`, which defaults to the current UTC year. When several archive files match a dataset and month, the most recently modified one takes precedence.

For each dataset and date, the priority is an explicitly supplied original report export, the CRM daily export, then the monthly archive. Provide `--legacy-source` or `UPB_LEGACY_SOURCE_DIR` to let the optional legacy daily files fill gaps when a monthly source cannot supply the date. The legacy folders are `UPB开卡数据/`, `UPB充值数据/`, and `UPB消费数据/{reap,passto,straitsx}/`. Legacy filenames consist of the export name followed by `YYYY-MM-DD.xlsx`. The **数据来源** (Sources) sheet records the actual files used.

### Running the report

Use Python 3.10+ and install openpyxl from `requirements.txt`. Excel generation also needs a Node.js runtime with access to `@oai/artifact-tool`. The tested setup uses the runtime bundled with Codex. For another setup that already has the library, set `UPB_NODE_RUNTIME` to the directory containing `bin/node` and the library's `node_modules`.

```sh
python3 -m pip install -r requirements.txt
python3 refresh.py --source "/path/to/upb-exports"
python3 refresh.py --history --source "/path/to/upb-exports" --legacy-source "/path/to/legacy-exports"
python3 refresh.py --history --from 2026-05-01 --to 2026-05-31 --source "/path/to/upb-exports"
python3 refresh.py --date 2026-05-12 --source "/path/to/upb-exports" --cards-export "/path/to/card-completions.xlsx"
```

Use `--output` to change the default `UPB每日数据汇总.xlsx` in the project directory. A single-date run also supports `--reap-export`, `--passto-export`, and `--straitsx-export`. The workbook records supplemental file paths and reuses those files for the same date while they are available.

The two `.command` launchers provide daily update and historical recalculation on macOS. Run `chmod +x *.command` if the downloaded launchers lack execute permissions. Double-clicking uses the default data directory and the Python visible to Finder. For a custom setup, use the terminal commands. Save and close the workbook before refreshing it.

### Implementation

1. **Select dates.** The daily run recomputes the two latest known UTC dates and appends available dates newer than the last workbook column. A first run starts with the two most recent available export dates. Use history mode or `--date` for older gaps. History mode rereads sources for the selected dates and defaults to ending yesterday in UTC.
2. **Read exports.** `refresh.py` uses openpyxl in read-only mode to find sheets with matching headers. Python reads separate files in parallel, parses amounts as `Decimal`, and checks file size and modification time after the read. Missing or invalid data is recorded for review.
3. **Aggregate and build Excel.** Python deduplicates and aggregates records by date, channel and segment. It passes daily summaries, provenance and manual inputs as transient JSON to `build.mjs`. The JavaScript builder uses Artifact Tool to produce the report and source sheets and adds Excel formulas for totals and shares.
4. **Save the result.** The existing workbook supplies manually entered registrations, user funds and investment balances, including explicit zero values. These follow their dates during a backfill. The tool locks the report, renders a candidate file in a temporary directory, checks that the original workbook has not been opened or changed, and replaces the workbook. Source files are read-only throughout.

### Calculation rules

| Metric | Rule |
| --- | --- |
| Card openings | Count records with a positive paid opening fee and a status of successful, in production or pending activation. Assign by completion date. For monthly archives with missing completion timestamps, use the application date. Deduplicate within date, channel and card type by card ID, falling back to order number. Daily application-time exports may omit cards requested earlier and completed on the report date. |
| Top-ups (USDT) | Sum the paid amount for successful top-up / collateral orders with USDT as the paid currency. Deduplicate by date and order number. Use the creation date for monthly archives and filename date for daily exports. Group by card channel, with remaining channels in Other. |
| Spending (USD) | Include debit rows. CRM daily, monthly and legacy files are limited to primary orders. Explicit original-report supplements retain their export scope. Reap uses the USD transaction amount. StraitsX uses the original transaction amount, or the transaction amount in the legacy schema, in USD. Passto converts its HKD transaction amount by dividing by the fixed **7.78** rate. Deduplicate by date and transaction ID. For conflicting versions in monthly archives, use the latest update when timestamps distinguish them. Monthly files are assigned by transaction date. |
| 2B / 2C | After trimming whitespace and ignoring case, an agent value equal to `UPay` is 2C. All other values are 2B. Top-ups use the assigned agent and transactions prefer the historical agent field for that channel. Each channel's 2B spending equals its total minus 2C spending. |
| Totals and shares | Round monetary subtotals to two decimals. Shares divide the segment or channel by its corresponding daily total. A zero denominator or missing required input yields a blank. |

Yellow source labels identify export scopes that warrant review. Missing exports, unsupported currencies, and unresolved duplicate conflicts leave the corresponding values blank and are recorded in the workbook. Export files, generated workbooks, and machine-specific configuration stay local.
