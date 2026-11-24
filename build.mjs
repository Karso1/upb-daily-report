import fs from "node:fs/promises";
import { createRequire } from "node:module";
import { homedir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const runtime = process.env.UPB_NODE_RUNTIME ?? join(homedir(), ".cache/codex-runtimes/codex-primary-runtime/dependencies/node");
const resolveRuntime = createRequire(join(runtime, "package.json"));
const { SpreadsheetFile, Workbook } = await import(pathToFileURL(resolveRuntime.resolve("@oai/artifact-tool")).href);

const [, , input, output, preview] = process.argv;
const records = JSON.parse(await fs.readFile(input, "utf8"));
const workbook = Workbook.create();
const summary = workbook.worksheets.add("UPB每日汇总");
const sources = workbook.worksheets.add("数据来源");
const navy = "#12323B";
const teal = "#08786C";
const pale = "#E8F5F2";
const amber = "#FFF0CE";
const channels = ["Reap", "Passto", "StraitsX"];
const rechargeChannels = [...channels, "其他渠道"];
const labels = [
  "UP Business 卡业务日汇总", "报表日期（UTC+0）", "平台注册人数", "开卡来源", "交易来源",
  "金融卡", "Reap 虚拟卡", "Reap 实体卡", "StraitsX 虚拟卡", "StraitsX 实体卡", "其他渠道开卡", "开卡汇总",
  "抵押 / 充值（USDT）", ...rechargeChannels, ...rechargeChannels.map(c => c + " 2C"), ...rechargeChannels.map(c => c + " 2B"),
  ...rechargeChannels.map(c => c + " 充值占比"), "2B 充值", "2C 充值", "2B 充值占比", "2C 充值占比", "充值汇总",
  "交易明细（USD）", ...channels, ...channels.map(c => c + " 2C"), ...channels.map(c => c + " 2B"),
  "2B 消费", "2C 消费", "2B 消费占比", "2C 消费占比", "消费汇总", "Passto 汇率（HKD/USD）",
  "资金明细（USD）", "用户资金", "理财金额",
];

function columnName(index) {
  let label = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    label = String.fromCharCode(65 + (n - 1) % 26) + label;
  }
  return label;
}

const lastColumn = columnName(records.length);
const summaryArea = summary.getRange(`A1:${lastColumn}${labels.length}`);
summaryArea.format.font.name = "Arial";
summaryArea.format.font.size = 11;
summaryArea.format.font.color = navy;
summaryArea.format.rowHeight = 21;
summaryArea.format.verticalAlignment = "center";
summary.getRange("A:A").format.columnWidthPx = 254;
summary.getRange(`B:${lastColumn}`).format.columnWidthPx = 156;
summary.getRange(`A1:A${labels.length}`).values = labels.map(label => [label]);
summary.showGridLines = false;
summary.freezePanes.freezeRows(2);
summary.freezePanes.freezeColumns(1);

for (const row of [1, 6, 13, 35, 51]) {
  const stripe = summary.getRange(`A${row}:${lastColumn}${row}`);
  stripe.format.fill = row === 1 ? navy : teal;
  stripe.format.font.color = "#FFFFFF";
  stripe.format.font.bold = true;
  stripe.format.rowHeight = row === 1 ? 35 : 24;
}
summary.getRange("A1").format.font.size = 16;
for (const row of [2, 12, 30, 31, 34, 45, 46, 49]) {
  const stripe = summary.getRange(`A${row}:${lastColumn}${row}`);
  stripe.format.fill = pale;
  stripe.format.font.bold = true;
}
for (const row of [3, 52, 53]) summary.getRange(`A${row}:${lastColumn}${row}`).format.fill = amber;
summary.getRange(`B4:${lastColumn}5`).format.font.size = 10;

records.forEach((record, index) => {
  const col = columnName(index + 1);
  const cell = row => summary.getRange(`${col}${row}`);
  const put = (row, value) => { cell(row).values = [[value]]; };
  const formula = (row, expression) => { cell(row).formulas = [[`=${expression.replaceAll("@", col)}`]]; };
  const subtotal = (row, start, end) => {
    const range = `@${start}:@${end}`;
    formula(row, `IF(COUNT(${range})=${end - start + 1},ROUND(SUM(${range}),2),"")`);
  };
  const share = (row, amount, total) => formula(row, `IF(AND(ISNUMBER(@${amount}),ISNUMBER(@${total}),@${total}>0),@${amount}/@${total},"")`);

  put(2, Date.parse(record.day + "T00:00:00Z") / 86400000 + 25569);
  put(3, record.registration);
  put(4, record.card_status);
  put(5, record.tx_status);
  if (record.card_status !== "完成日导出") cell(4).format.fill = amber;
  if (record.tx_status !== "原日报导出") cell(5).format.fill = amber;

  summary.getRange(`${col}7:${col}11`).values = record.cards.map(value => [value]);
  subtotal(12, 7, 11);
  summary.getRange(`${col}18:${col}25`).values = record.recharge.map(value => [value]);
  for (let i = 0; i < rechargeChannels.length; i++) {
    formula(14 + i, `IF(COUNT(@${18 + i},@${22 + i})=2,ROUND(SUM(@${18 + i},@${22 + i}),2),"")`);
    share(26 + i, 14 + i, 34);
  }
  subtotal(30, 22, 25);
  subtotal(31, 18, 21);
  share(32, 30, 34);
  share(33, 31, 34);
  subtotal(34, 30, 31);
  for (let i = 0; i < channels.length; i++) {
    const [total, consumer] = record.transactions[i];
    put(36 + i, total);
    put(39 + i, consumer);
    formula(42 + i, `IF(COUNT(@${36 + i},@${39 + i})=2,ROUND(@${36 + i}-@${39 + i},2),"")`);
  }
  subtotal(45, 42, 44);
  subtotal(46, 39, 41);
  share(47, 45, 49);
  share(48, 46, 49);
  subtotal(49, 36, 38);
  if (record.transactions[1][0] !== null) put(50, 7.78);
  put(52, record.user_funds);
  put(53, record.wealth);
});

for (const [rows, mask] of [
  [[2], "yyyy-mm-dd"], [[3, 7, 8, 9, 10, 11, 12], "#,##0"],
  [[14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 30, 31, 34, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 49, 52, 53], "#,##0.00"],
  [[26, 27, 28, 29, 32, 33, 47, 48], "0.00%"], [[50], "0.00"],
]) {
  for (const row of rows) summary.getRange(`B${row}:${lastColumn}${row}`).setNumberFormat(mask);
}

sources.showGridLines = false;
sources.freezePanes.freezeRows(2);
sources.getRange("A1:H1").merge();
sources.getRange("A1").values = [["UPB 日报数据来源"]];
sources.getRange("A2:H2").values = [["日期", "开卡", "充值", ...channels, "补充导出", "待核对明细"]];
sources.getRange(`A3:H${records.length + 2}`).values = records.map(record => [
  record.day, record.paths.cards, record.paths.recharge, ...channels.map(channel => record.paths[channel]),
  record.supplemental.join(","), record.issues,
]);
sources.getRange("A:A").format.columnWidthPx = 108;
sources.getRange("B:F").format.columnWidthPx = 220;
sources.getRange("G:G").format.columnWidthPx = 172;
sources.getRange("H:H").format.columnWidthPx = 300;
const sourceArea = sources.getRange(`A1:H${records.length + 2}`);
sourceArea.format.font.name = "Arial";
sourceArea.format.font.size = 10;
sourceArea.format.font.color = navy;
sourceArea.format.verticalAlignment = "center";
sources.getRange("A1:H2").format.fill = navy;
sources.getRange("A1:H2").format.font.color = "#FFFFFF";
sources.getRange("A1:H2").format.font.bold = true;
sources.getRange("A1:H2").format.rowHeight = 25;
const fileArea = sources.getRange(`B3:H${records.length + 2}`);
fileArea.format.wrapText = true;
fileArea.format.rowHeight = 60;

const notes = [
  "平台注册人数、用户资金、理财金额为黄色手填行，按日期保留已保存的数字。",
  "开卡只计实付费大于 0 的开卡成功、制卡中、待激活。按完成日归集，旧月档缺完成时间时按申请日归集。",
  "充值取充值/抵押成功订单的 USDT 实付金额。充值按所属代理商、交易按渠道导出中的历史代理商分组，UPay 归 2C，其余归 2B。",
  "交易按支出汇总，日报补充导出沿用其原有范围。CRM 日档和归档仅统计主订单。Passto 采用 7.78 HKD/USD。",
  "历史计算优先用补充导出、CRM 日档、月档，旧版日档补月档漏缺。月档充值按创建日、交易按交易日归集。",
  "历史开卡和充值中的 OE 等渠道归入其他渠道。月档中同一交易 ID 有多次更新时使用较新的记录。",
  "缺少源表或相关日期的明细时对应金额或张数留空。黄色来源标记表示导出范围待核对。",
];
notes.forEach((note, index) => {
  const row = records.length + 4 + index;
  const line = sources.getRange(`A${row}:H${row}`);
  line.merge();
  sources.getRange(`A${row}`).values = [[note]];
  line.format.font.name = "Arial";
  line.format.font.color = navy;
  line.format.rowHeight = 22;
});

workbook.recalculate();
if (preview) {
  const lastPreviewColumn = columnName(Math.min(records.length, 4));
  const blob = await workbook.render({ sheetName: summary.name, range: `A1:${lastPreviewColumn}${labels.length}`, format: "png" });
  await fs.writeFile(preview, new Uint8Array(await blob.arrayBuffer()));
  const sourceBlob = await workbook.render({ sheetName: sources.name, range: `A1:H${Math.min(records.length + 10, 13)}`, format: "png" });
  await fs.writeFile(preview.replace(/\.png$/, "-sources.png"), new Uint8Array(await sourceBlob.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(output);
