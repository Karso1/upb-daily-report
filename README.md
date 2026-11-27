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
| PhotonPay 交易 | 光子易渠道明细账务 | `消费数据/每日自动下载/YYYY-MM/光子易_YYYY-MM-DD.xlsx` 或同目录下的 `PhotonPay_YYYY-MM-DD.xlsx` |

开卡的 CRM 日档按申请时间导出，充值日档按订单创建时间导出。Passto、Reap 和 PhotonPay 的 CRM 交易日档按更新时间导出，StraitsX 日档按交易时间导出。交易日档按整个文件归到文件名对应的日期，因此更新时间口径的日档可能含有更早发生的交易。PhotonPay 自动日档缺失时，还会在 `消费数据/历史补充/` 下查找同日的 `光子易_YYYY-MM-DD.xlsx` 或 `PhotonPay_YYYY-MM-DD.xlsx`。

历史模式还识别开卡目录下带年月的归档、`充值数据/Q*/` 下带年月的充值月档，以及 `消费数据/UPB/Q*/<月份>/` 下以渠道名开头的交易月档。交易月档文件名无年份时，使用 `--archive-year` 指定的年份，默认当年 UTC 年份。一个渠道同月有多份候选月档时，选择最后修改的文件。

来源按日期、按指标选择。优先级为指定的原日报补充导出、CRM 日档、月档。配置 `--legacy-source` 或 `UPB_LEGACY_SOURCE_DIR` 后，月档缺少相应日期或无法计算时，还会尝试旧版每日数据库。旧日档使用 `UPB开卡数据/`、`UPB充值数据/`、`UPB消费数据/{reap,passto,straitsx}/`，文件名为导出名紧跟 `YYYY-MM-DD.xlsx`。每一日实际选用的文件在 Excel 的「数据来源」页可查。

### 使用方法

需要 **Python 3.10+**。下载或克隆仓库，在项目目录中安装 `requirements.txt`。openpyxl 负责读取 Excel，XlsxWriter 负责生成 Excel。运行所需的依赖都可从 PyPI 安装。

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

上面的安装命令适用于 macOS 和 Linux。Windows 可运行 `py -3 -m venv .venv`，再用 `.venv\Scripts\python.exe -m pip install -r requirements.txt`。以下命令中的 `.venv/bin/python` 在 Windows 换成 `.venv\Scripts\python.exe`。

```sh
# 日常更新
.venv/bin/python refresh.py --source "/path/to/upb-exports"

# 重算所有可识别的历史日期
.venv/bin/python refresh.py --history --source "/path/to/upb-exports" --legacy-source "/path/to/legacy-exports"

# 重算一个区间，含起止日期
.venv/bin/python refresh.py --history --from 2026-05-01 --to 2026-05-31 --source "/path/to/upb-exports"

# 重算单日，指定该日按完成时间导出的开卡明细
.venv/bin/python refresh.py --date 2026-05-12 --source "/path/to/upb-exports" --cards-export "/path/to/card-completions.xlsx"
```

默认输出为项目目录下的 `UPB每日数据汇总.xlsx`，可通过 `--output` 指定。交易补充文件分别通过 `--reap-export`、`--passto-export`、`--straitsx-export` 和 `--photonpay-export` 指定，须与 `--date` 一起使用。补充文件路径会写进来源页，文件仍在时，重算同一天会复用它。

macOS 安装好依赖后，把导出目录放在项目的 `data/`，也可以在项目目录建立指向现有导出目录的链接 `ln -s "/path/to/upb-exports" data`。首次使用时运行 `chmod +x *.command`，然后在 Finder 双击 `双击更新UPB日报.command` 或 `双击重算UPB历史.command`。入口会优先使用项目的 `.venv/bin/python`，执行成功后打开汇总表。指定其他输入或输出时可用上面的终端命令。更新前先保存并关闭 Excel。

### 实现步骤和技术

1. **确定日期。** 日常模式默认重算最近两个已知 UTC 日期，并追加汇总末尾之后已下载的日期。首次运行默认取有源文件的最近两个日期。中间漏掉的旧日期用历史模式或 `--date` 补。历史模式重新读取所选日期的归档和日档，默认截至 UTC 昨天。
2. **读取源表。** `refresh.py` 使用 openpyxl 的只读模式，查找表头符合要求的工作表。Python 按文件并行读取，解析金额时使用 `Decimal`。读完检查源文件大小和修改时间，遇到来源缺失或数据问题时记录待核对信息。
3. **汇总并生成 Excel。** Python 按渠道和客群归集并去重，把各日期的汇总、来源和手填值交给 `report_workbook.py`。XlsxWriter 生成「UPB每日汇总」和「数据来源」，写入合计和占比公式，同时保存初始计算结果，方便 Excel 和读取缓存值的工具直接查看。
4. **保存。** 更新时从现有日报按日期读回「平台注册人数」「用户资金」「理财金额」，包括手填的 0。已存在的日期原位更新。程序对同一个输出加锁，在临时目录生成新工作簿，再检查旧文件是否被打开或修改，最后替换输出。源 Excel 全程只读。

### 指标怎么算

| 指标 | 规则 |
| --- | --- |
| 开卡 | 只计实付开卡费大于 0，且状态为「开卡成功」「制卡中」或「待激活」的记录。优先按开卡成功时间归日，月档该字段为空时用申请日。按日期、渠道和卡种用卡 ID 去重，缺卡 ID 时用订单号。单份日档若按申请时间导出，可能漏掉早前申请、当天才完成的卡。 |
| 充值（USDT） | 只计状态为「充值/抵押成功」的订单，累加 `实付金额`，要求实付币种为 USDT，按日期和订单号去重。月档按创建时间归日，日档按文件名日期归日。通过 `卡片渠道` 分组，OE 单列。未单列的渠道有充值金额时显示在「其他渠道」。 |
| 消费（USD） | 只计「支出」。CRM 日档、月档及旧版日档限「主订单」，逐日指定的原日报补充文件沿用其筛选范围。Reap 读 `交易金额`（USD），StraitsX 优先读 `交易原始金额`（旧版用 `交易金额`，USD），Passto 读 HKD `交易金额` 后除以固定的 `7.78`。按日期和交易 ID 去重，月档同一 ID 有金额或归属变动且更新时间可区分时取较新的记录。月档按交易时间归日。 |
| 2B／2C | 代理商字段去掉前后空白并忽略大小写后，等于 `UPay` 的归 2C，其余归 2B。充值读取所属代理商，交易优先读取渠道的历史代理商字段。交易的 2B 金额为渠道合计减渠道 2C。 |
| 合计和占比 | 金额按分组小计四舍五入至两位小数。占比为分组或渠道金额除以同日对应总额。分母为零或有必要输入缺失时，留空。 |

#### PhotonPay 的 2B／2C

- **充值（USDT）**。充值表的「卡片渠道」为 `photonpay`（忽略大小写）或「光子易」时，统一归入 PhotonPay。只计「充值/抵押成功」且实付币种为 USDT 的订单，按订单号和日期去重。根据「所属代理商」拆分后，分别累加 2B、2C 的「实付金额」。旧表缺该列时使用「历史代理商」。
- **消费（USD）**。光子易或 PhotonPay 交易表只计「主订单」中的「支出」，读取 USD「交易金额」，按日期和交易 ID 去重，支持导出列「主ID」。优先用「历史所属代理商」区分客群，缺该列时依次使用「历史代理商名称」和「所属代理商」。PhotonPay 2C 为 UPay 记录的金额合计，PhotonPay 2B 为渠道合计减去 2C。
- **合计与缺档**。PhotonPay 充值和消费均计入日报对应合计。首次有 PhotonPay 源表的日期之前，消费合计沿用原三渠道。从该日期开始，缺少 PhotonPay 源表会标记待核对，消费合计和占比留空。已有源表但无符合条件的支出时，PhotonPay 消费计为零。

#### OE 充值

「卡片渠道」为 `OE` 的成功 USDT 充值单列，识别时去掉前后空白并忽略大小写。OE 2C 累计所属代理商为 `UPay` 的实付金额，OE 2B 累计其余代理商的实付金额。OE 充值占比等于 OE 充值合计除以同日全部渠道的充值汇总。

旧工作簿里的「其他渠道」可能包含多个渠道。升级时保留原分组金额和手填值。用历史重算读取原始导出后，确认属于 OE 的订单会计入 OE，其余继续计入「其他渠道」。

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
| PhotonPay transactions | 光子易 channel ledger | `消费数据/每日自动下载/YYYY-MM/光子易_YYYY-MM-DD.xlsx` or `PhotonPay_YYYY-MM-DD.xlsx` in the same folder |

The upstream CRM card export is filtered by application time, and the top-up export by creation time. The Passto, Reap, and PhotonPay daily exports use update time, while StraitsX uses transaction time. The report assigns each daily transaction file to its filename date, so an update-time batch can contain transactions that took place earlier. When the PhotonPay daily file is missing, the report also looks under `消费数据/历史补充/` for an export named `光子易_YYYY-MM-DD.xlsx` or `PhotonPay_YYYY-MM-DD.xlsx` for that date.

History mode also searches for month-dated card files under `开卡/`, month-dated top-ups under `充值数据/Q*/`, and channel-named ledgers under `消费数据/UPB/Q*/<month>/`. For transaction archive filenames without a year, set `--archive-year`, which defaults to the current UTC year. When several archive files match a dataset and month, the most recently modified one takes precedence.

For each dataset and date, the priority is an explicitly supplied original report export, the CRM daily export, then the monthly archive. Provide `--legacy-source` or `UPB_LEGACY_SOURCE_DIR` to let the optional legacy daily files fill gaps when a monthly source cannot supply the date. The legacy folders are `UPB开卡数据/`, `UPB充值数据/`, and `UPB消费数据/{reap,passto,straitsx}/`. Legacy filenames consist of the export name followed by `YYYY-MM-DD.xlsx`. The **数据来源** (Sources) sheet records the actual files used.

### Running the report

Use **Python 3.10+**. Download or clone the repository, then install `requirements.txt` from the project directory. openpyxl reads the source workbooks and XlsxWriter generates the report. All runtime dependencies are available from PyPI.

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

These setup commands are for macOS and Linux. On Windows, use `py -3 -m venv .venv` and `.venv\Scripts\python.exe -m pip install -r requirements.txt`. Replace `.venv/bin/python` in the examples with `.venv\Scripts\python.exe` on Windows.

```sh
.venv/bin/python refresh.py --source "/path/to/upb-exports"
.venv/bin/python refresh.py --history --source "/path/to/upb-exports" --legacy-source "/path/to/legacy-exports"
.venv/bin/python refresh.py --history --from 2026-05-01 --to 2026-05-31 --source "/path/to/upb-exports"
.venv/bin/python refresh.py --date 2026-05-12 --source "/path/to/upb-exports" --cards-export "/path/to/card-completions.xlsx"
```

Use `--output` to change the default `UPB每日数据汇总.xlsx` in the project directory. A single-date run also supports `--reap-export`, `--passto-export`, `--straitsx-export`, and `--photonpay-export`. The workbook records supplemental file paths and reuses those files for the same date while they are available.

On macOS, put the exports in the project's `data/` directory or link it to your existing export folder with `ln -s "/path/to/upb-exports" data`. Run `chmod +x *.command` during setup, then double-click `双击更新UPB日报.command` (daily update) or `双击重算UPB历史.command` (historical recalculation) in Finder. The launchers prefer the project's `.venv/bin/python` and open the workbook on success. Use the terminal for custom input or output paths. Save and close the workbook before refreshing it.

### Implementation

1. **Select dates.** The daily run recomputes the two latest known UTC dates and appends available dates newer than the last workbook column. A first run starts with the two most recent available export dates. Use history mode or `--date` for older gaps. History mode rereads sources for the selected dates and defaults to ending yesterday in UTC.
2. **Read exports.** `refresh.py` uses openpyxl in read-only mode to find sheets with matching headers. Python reads separate files in parallel, parses amounts as `Decimal`, and checks file size and modification time after the read. Missing or invalid data is recorded for review.
3. **Aggregate and build Excel.** Python deduplicates and aggregates records by date, channel and segment. It passes daily summaries, provenance and manual inputs to `report_workbook.py`. XlsxWriter creates the report and source sheets, writes Excel formulas for totals and shares, and stores their initial cached results.
4. **Save the result.** The existing workbook supplies manually entered registrations, user funds and investment balances, including explicit zero values. These follow their dates during a backfill. The tool locks the report, renders a candidate file in a temporary directory, checks that the original workbook has not been opened or changed, and replaces the workbook. Source files are read-only throughout.

### Calculation rules

| Metric | Rule |
| --- | --- |
| Card openings | Count records with a positive paid opening fee and a status of successful, in production or pending activation. Assign by completion date. For monthly archives with missing completion timestamps, use the application date. Deduplicate within date, channel and card type by card ID, falling back to order number. Daily application-time exports may omit cards requested earlier and completed on the report date. |
| Top-ups (USDT) | Sum the paid amount for successful top-up / collateral orders with USDT as the paid currency. Deduplicate by date and order number. Use the creation date for monthly archives and filename date for daily exports. Group by card channel. OE is listed separately. Channels without a dedicated row appear in Other when they have top-up amounts. |
| Spending (USD) | Include debit rows. CRM daily, monthly and legacy files are limited to primary orders. Explicit original-report supplements retain their export scope. Reap uses the USD transaction amount. StraitsX uses the original transaction amount, or the transaction amount in the legacy schema, in USD. Passto converts its HKD transaction amount by dividing by the fixed **7.78** rate. Deduplicate by date and transaction ID. For conflicting versions in monthly archives, use the latest update when timestamps distinguish them. Monthly files are assigned by transaction date. |
| 2B / 2C | After trimming whitespace and ignoring case, an agent value equal to `UPay` is 2C. All other values are 2B. Top-ups use the assigned agent and transactions prefer the historical agent field for that channel. Each channel's 2B spending equals its total minus 2C spending. |
| Totals and shares | Round monetary subtotals to two decimals. Shares divide the segment or channel by its corresponding daily total. A zero denominator or missing required input yields a blank. |

#### PhotonPay 2B / 2C

- **Top-ups (USDT).** Card channel values of `photonpay` (case-insensitive) or `光子易` map to PhotonPay. Keep successful top-up or collateral orders paid in USDT and deduplicate by date and order number. Split records by assigned agent (`所属代理商`) and sum the paid amount for each segment. If a legacy sheet lacks that column, use historical agent (`历史代理商`).
- **Spending (USD).** Keep primary orders with debit flow. Sum the USD transaction amount (`交易金额`) and deduplicate by date and transaction ID, including exports that label the ID `主ID`. Use historical assigned agent (`历史所属代理商`), falling back to `历史代理商名称` and then `所属代理商` when the preceding column is absent. PhotonPay 2C is the sum of UPay transactions. PhotonPay 2B is the channel total minus 2C.
- **Totals and missing files.** PhotonPay contributes to both daily totals. Spending totals use the original three channels for dates before the first available PhotonPay export. From that date onward, a missing PhotonPay file flags the date for review and leaves spending totals and shares blank. A present file with no qualifying debits contributes zero.

#### OE top-ups

Successful USDT top-ups with a card channel of `OE` have their own row. The channel name is trimmed and matched case-insensitively. OE 2C sums the paid amounts assigned to `UPay`, and OE 2B sums the paid amounts assigned to all other agents. OE's share is its top-up total divided by the daily total across all channels.

Older workbooks may have several channels grouped under Other. An upgrade preserves that amount and the manual inputs. When history is recalculated from the original exports, orders identified as OE move to OE and the remaining orders stay in Other.

Yellow source labels identify export scopes that warrant review. Missing exports, unsupported currencies, and unresolved duplicate conflicts leave the corresponding values blank and are recorded in the workbook. Export files, generated workbooks, and machine-specific configuration stay local.
