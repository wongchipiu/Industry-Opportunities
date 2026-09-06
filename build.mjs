import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const outputDir = path.join(repoRoot, "outputs/01a05331-247c-7c81-86f0-48c6c2a5ce93");
const outputPath = `${outputDir}/印尼高增长行业机会_2026.xlsx`;

const wb = Workbook.create();
const summary = wb.worksheets.add("Summary");
const score = wb.worksheets.add("Scorecard");
const unit = wb.worksheets.add("Unit Economics");
const plan = wb.worksheets.add("90-Day Plan");
const regulatory = wb.worksheets.add("Regulatory");
const sources = wb.worksheets.add("Sources");
const checks = wb.worksheets.add("Checks");

wb.comments.setSelf({ displayName: "Bruce Huang" });

const navy = "#17365D";
const teal = "#0F766E";
const blue = "#DCE6F1";
const paleBlue = "#EAF2F8";
const paleGreen = "#E2F0D9";
const paleYellow = "#FFF2CC";
const paleRed = "#FCE4D6";
const grey = "#F2F2F2";
const white = "#FFFFFF";
const dark = "#1F2937";
const inputBlue = "#0000FF";
const linkGreen = "#008000";

function titleBand(sheet, range, text) {
  sheet.getRange(range).merge();
  const cell = sheet.getRange(range.split(":")[0]);
  cell.values = [[text]];
  cell.format = {
    fill: navy,
    font: { bold: true, color: white, size: 18 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
  sheet.getRange(range).format.rowHeight = 32;
}

function sectionBand(sheet, range, text) {
  sheet.getRange(range).merge();
  const cell = sheet.getRange(range.split(":")[0]);
  cell.values = [[text]];
  cell.format = {
    fill: navy,
    font: { bold: true, color: white, size: 11 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
}

function styleHeader(range) {
  range.format = {
    fill: teal,
    font: { bold: true, color: white },
    verticalAlignment: "center",
    horizontalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: "#B7C9D6" },
  };
}

function styleInput(range, numberFormat) {
  range.format.fill = paleYellow;
  range.format.font = { color: inputBlue };
  if (numberFormat) range.format.numberFormat = numberFormat;
}

for (const sheet of [summary, score, unit, plan, regulatory, sources, checks]) {
  sheet.showGridLines = false;
}

// Sources and external inputs
titleBand(sources, "A1:H1", "印尼机会研究 — 数据来源与外部输入");
sources.getRange("A2:H2").merge();
sources.getRange("A2").values = [["截至 2026-08-30。数值均保留原始单位；URL 为可审计来源。蓝字/黄底为用户可修改假设，绿色字为跨表引用。"]];
sources.getRange("A2").format = { font: { color: "#666666", italic: true }, wrapText: true };
sources.getRange("A4:H4").values = [["ID", "指标", "数值", "单位", "期间/日期", "来源", "URL", "说明"]];
styleHeader(sources.getRange("A4:H4"));
const sourceRows = [
  ["S01", "印尼实际GDP同比", 5.29, "%", "2026 Q2", "BPS", "https://www.bps.go.id/en/pressrelease/2026/08/05/2605/ekonomi-indonesia-triwulan-ii-2026-tumbuh-5-29-persen--y-on-y-.html", "生产法口径；二季度同比"],
  ["S02", "住宿餐饮业同比", 11.83, "%", "2026 H1", "BPS", "https://www.bps.go.id/en/pressrelease/2026/08/05/2605/ekonomi-indonesia-triwulan-ii-2026-tumbuh-5-29-persen--y-on-y-.html", "上半年生产侧增速最高"],
  ["S03", "2025年进口总额", 241362.6, "USD mn", "2025", "BPS", "https://www.bps.go.id/en/publication/2026/08/04/29900d4157c9ba7ad7fc72ea/statistik-perdagangan-luar-negeri-indonesia-impor-2025.html", "同比 +2.62%"],
  ["S04", "中国占印尼进口比重", 36.22, "%", "2025", "BPS", "https://www.bps.go.id/en/publication/2026/08/04/29900d4157c9ba7ad7fc72ea/statistik-perdagangan-luar-negeri-indonesia-impor-2025.html", "中国进口额 USD87.41bn"],
  ["S05", "资本品占进口比重", 20.6, "%", "2025", "BPS", "https://www.bps.go.id/en/publication/2026/08/04/29900d4157c9ba7ad7fc72ea/statistik-perdagangan-luar-negeri-indonesia-impor-2025.html", "资本品 USD49.72bn"],
  ["S06", "下游产业投资同比", 43.3, "%", "2025", "BKPM", "https://www.bkpm.go.id/en/info/press-release/investment-realization-in-2025-exceeds-target-downstream-investment-jumps-43-3-percent", "下游产业投资 IDR584.1tn"],
  ["S07", "电商FMCG销售额", 88.4, "IDR tn", "2026 H1", "Compas", "https://compas.co.id/article/fmcg-e-commerce-tembus-88trilliun/", "同比约 +29%；覆盖主要平台"],
  ["S08", "美妆个护电商销售额", 42.9, "IDR tn", "2026 H1", "Compas/Katadata", "https://databoks.katadata.co.id/en/technology-telecommunications/statistics/6a8e9b010d73f/beauty-products-lead-fmcg-sales-in-indonesian-e-commerce-h1-2026", "占电商FMCG 48.6%"],
  ["S09", "视频电商交易量同比", 90, "%", "2025", "Google/Temasek/Bain", "https://services.google.com/fh/files/misc/indonesia_e_conomy_sea_2025_report.pdf", "年交易量 26 亿笔"],
  ["S10", "视频电商典型客单价", 5.25, "USD/order", "2025", "Google/Temasek/Bain", "https://services.google.com/fh/files/misc/indonesia_e_conomy_sea_2025_report.pdf", "报告给出的区间为 USD4.5–6"],
  ["S11", "CNY/IDR参考汇率", 2634.06, "IDR/CNY", "2026-08-28", "Bank Indonesia", "https://www.bi.go.id/en/statistik/informasi-kurs/kurs-acuan-non-usd-idr/default.aspx", "用于预算换算；请按实际付款日更新"],
  ["S12", "PT PMA最低实缴资本", 2500, "IDR mn", "2025-10-02", "UNCTAD/BKPM Reg. 5/2025", "https://investmentpolicy.unctad.org/investment-policy-monitor/measures/5131/indonesia-lowers-paid-up-capital-requirement-for-foreign-owned-limited-liability-companies", "一般投资计划仍需 >IDR10bn/KBLI/地点"],
  ["S13", "RUPTL新增可再生能源", 42.6, "GW", "2025–2034", "ESDM", "https://www.esdm.go.id/id/media-center/arsip-berita/menteri-esdm-umumkan-ruptl-pln-2025-2034-serap-lebih-dari-17-juta-tenaga-kerja-baru", "另含 10.3GW 储能"],
];
sources.getRange(`A5:H${4 + sourceRows.length}`).values = sourceRows;
sources.getRange("C5:C17").format.numberFormat = "#,##0.00";
sources.getRange("A5:H17").format.verticalAlignment = "top";
sources.getRange("A5:H17").format.rowHeight = 40;
sources.getRange("B5:B17").format.wrapText = true;
sources.getRange("F5:H17").format.wrapText = true;
sources.getRange("G5:G17").format.font = { color: "#0563C1", underline: true };
sources.freezePanes.freezeRows(4);
sources.getRange("A:A").format.columnWidth = 8;
sources.getRange("B:B").format.columnWidth = 28;
sources.getRange("C:C").format.columnWidth = 14;
sources.getRange("D:E").format.columnWidth = 14;
sources.getRange("F:F").format.columnWidth = 24;
sources.getRange("G:G").format.columnWidth = 55;
sources.getRange("H:H").format.columnWidth = 34;
for (let r = 5; r <= 17; r++) {
  const url = sourceRows[r - 5][6];
  wb.comments.addThread({ cell: sources.getRange(`C${r}`) }, `Source: ${url}`);
}

// Scorecard
titleBand(score, "A1:M1", "候选赛道量化评分");
score.getRange("A2:M2").merge();
score.getRange("A2").values = [["评分 1–10，越高越好。监管难度和竞争强度已经反向处理：高分代表更容易进入/竞争更友好。权重可修改，综合分自动更新。"]];
score.getRange("A2").format = { font: { color: "#666666", italic: true }, wrapText: true };
score.getRange("A3:B3").values = [["权重", ""]];
score.getRange("C3:J3").values = [[0.20, 0.15, 0.10, 0.10, 0.10, 0.10, 0.15, 0.10]];
score.getRange("C3:J3").format.numberFormat = "0%";
styleInput(score.getRange("C3:J3"), "0%");
score.getRange("A4:M4").values = [["排名", "赛道", "需求增长", "中国供应链优势", "毛利潜力", "复购/持续需求", "渠道可达性", "监管便利", "资本效率", "竞争友好度", "综合分", "关键证据", "建议切口"]];
styleHeader(score.getRange("A4:M4"));
const scoreRows = [
  [null, "工业MRO/自动化/节能", 8.5, 9.0, 7.5, 9.0, 6.5, 7.0, 8.0, 7.5, null, "下游投资+43.3%；资本品进口占20.6%", "工厂备件仓＋售后工程师；先做泵阀/传感器/过滤密封"],
  [null, "摩托车/汽车后市场", 8.0, 8.5, 7.0, 8.0, 8.0, 6.5, 8.0, 6.5, null, "2026年零配件零售持续正增长", "骑手耐用配件、维修工具、诊断设备；避开假货和低价件"],
  [null, "本地化美妆个护", 8.5, 7.5, 8.5, 8.0, 9.0, 4.5, 7.0, 5.0, null, "H1 2026线上销售IDR42.9tn；视频电商美妆占20%", "3–5个高演示性SKU；本地OEM/进口商＋达人联盟"],
  [null, "家清浓缩/防虫", 9.0, 7.5, 7.0, 9.0, 8.5, 4.5, 6.5, 5.0, null, "H1 2026 Homecare为线上增速最快品类之一", "浓缩补充装＋本地灌装；先避开高监管杀虫剂"],
  [null, "冷链/餐饮/酒店设备", 8.5, 8.0, 7.0, 7.5, 6.5, 6.5, 6.0, 7.0, null, "H1 2026住宿餐饮+11.83%", "制冰机、商用冷柜、温控、节能与维护合同"],
  [null, "工商业光伏/储能服务", 8.0, 9.0, 7.5, 6.5, 4.5, 4.5, 3.5, 6.5, null, "RUPTL计划42.6GW可再生能源＋10.3GW储能", "先做EMS/BMS、运维、消防热管理，不先做重资产EPC"],
];
score.getRange("A5:M10").values = scoreRows;
score.getRange("K5").formulas = [["=SUMPRODUCT(C5:J5,$C$3:$J$3)"]];
score.getRange("K5:K10").fillDown();
score.getRange("A5").formulas = [["=RANK(K5,$K$5:$K$10,0)"]];
score.getRange("A5:A10").fillDown();
score.getRange("C5:J10").format.numberFormat = "0.0";
styleInput(score.getRange("C5:J10"), "0.0");
score.getRange("K5:K10").format.numberFormat = "0.00";
score.getRange("K5:K10").format.font = { bold: true };
score.getRange("K5:K10").conditionalFormats.add("colorScale", {
  colors: ["#F8696B", "#FFEB84", "#63BE7B"],
  thresholds: ["min", "50%", "max"],
});
score.getRange("L5:M10").format.wrapText = true;
score.getRange("A5:M10").format.verticalAlignment = "top";
score.getRange("A5:M10").format.rowHeight = 48;
score.getRange("A2:M2").format.rowHeight = 28;
score.freezePanes.freezeRows(4);
score.freezePanes.freezeColumns(2);
score.getRange("A:A").format.columnWidth = 8;
score.getRange("B:B").format.columnWidth = 25;
score.getRange("C:J").format.columnWidth = 13;
score.getRange("K:K").format.columnWidth = 12;
score.getRange("L:L").format.columnWidth = 34;
score.getRange("M:M").format.columnWidth = 48;
for (let r = 5; r <= 10; r++) {
  wb.comments.addThread({ cell: score.getRange(`K${r}`) }, "Analyst score based on demand evidence, China supply advantage, route-to-market, regulation, capital needs and competition. Edit the eight blue input scores or weights to run your own view.");
}

// Unit economics
titleBand(unit, "A1:D1", "三条优先赛道 — 单位经济模型（人民币）");
unit.getRange("A2:D2").values = [["指标", "工业MRO", "本地化美妆", "摩托后市场"]];
styleHeader(unit.getRange("A2:D2"));
const unitLabels = [
  "稳态月订单数", "平均客单价（CNY）", "月含税成交额", "取消/退货率", "月净收入",
  "采购/生产成本率", "国际物流关税率", "平台/支付费率", "销售/达人佣金率", "广告率", "本地履约率", "质保/售后率",
  "总变动成本率", "月贡献毛利", "贡献毛利率", "月固定运营费", "月EBITDA", "盈亏平衡订单数",
  "试点首月订单", "订单月增速", "首年累计净收入", "首年累计EBITDA"
];
unit.getRange("A3:A24").values = unitLabels.map(x => [x]);
const inputRows = [
  [20, 6000, 2500],
  [18000, 75, 180],
  [null, null, null],
  [0.02, 0.05, 0.04],
  [null, null, null],
  [0.55, 0.25, 0.46],
  [0.08, 0.10, 0.11],
  [0.01, 0.09, 0.08],
  [0.06, 0.15, 0.06],
  [0.03, 0.10, 0.06],
  [0.02, 0.08, 0.06],
  [0.02, 0.03, 0.03],
  [null, null, null],
  [null, null, null],
  [null, null, null],
  [80000, 90000, 70000],
  [null, null, null],
  [null, null, null],
  [5, 800, 350],
  [0.18, 0.16, 0.15],
  [null, null, null],
  [null, null, null],
];
unit.getRange("B3:D24").values = inputRows;
for (const col of ["B", "C", "D"]) {
  unit.getRange(`${col}5`).formulas = [[`=ROUND(${col}3*${col}4,0)`]];
  unit.getRange(`${col}7`).formulas = [[`=ROUND(${col}5*(1-${col}6),0)`]];
  unit.getRange(`${col}15`).formulas = [[`=ROUND(SUM(${col}8:${col}14),4)`]];
  unit.getRange(`${col}16`).formulas = [[`=ROUND(${col}7*(1-${col}15),0)`]];
  unit.getRange(`${col}17`).formulas = [[`=ROUND(${col}16/${col}7,4)`]];
  unit.getRange(`${col}19`).formulas = [[`=ROUND(${col}16-${col}18,0)`]];
  unit.getRange(`${col}20`).formulas = [[`=ROUND(${col}18/(${col}4*(1-${col}6)*${col}17),0)`]];
}
styleInput(unit.getRange("B3:D4"), "#,##0");
styleInput(unit.getRange("B6:D6"), "0.0%");
styleInput(unit.getRange("B8:D14"), "0.0%");
styleInput(unit.getRange("B18:D18"), "¥#,##0;[Red](¥#,##0);-");
styleInput(unit.getRange("B21:D22"), "0.0%");
unit.getRange("B21:D21").format.numberFormat = "#,##0";
unit.getRange("B3:D5").format.numberFormat = "#,##0";
unit.getRange("B7:D7").format.numberFormat = "¥#,##0;[Red](¥#,##0);-";
unit.getRange("B8:D15").format.numberFormat = "0.0%";
unit.getRange("B16:D16").format.numberFormat = "¥#,##0;[Red](¥#,##0);-";
unit.getRange("B17:D17").format.numberFormat = "0.0%";
unit.getRange("B19:D19").format.numberFormat = "¥#,##0;[Red](¥#,##0);-";
unit.getRange("B20:D20").format.numberFormat = "#,##0";
unit.getRange("B23:D24").format.numberFormat = "¥#,##0;[Red](¥#,##0);-";
sectionBand(unit, "A27:J27", "12个月爬坡情景");
unit.getRange("A28:J28").values = [["月", "MRO订单", "MRO净收入", "MRO EBITDA", "美妆订单", "美妆净收入", "美妆 EBITDA", "摩托订单", "摩托净收入", "摩托 EBITDA"]];
styleHeader(unit.getRange("A28:J28"));
for (let i = 0; i < 12; i++) {
  const r = 29 + i;
  unit.getRange(`A${r}`).values = [[i + 1]];
  unit.getRange(`B${r}`).formulas = [[`=ROUND($B$21*(1+$B$22)^(A${r}-1),0)`]];
  unit.getRange(`C${r}`).formulas = [[`=ROUND(B${r}*$B$4*(1-$B$6),0)`]];
  unit.getRange(`D${r}`).formulas = [[`=ROUND(C${r}*$B$17-$B$18,0)`]];
  unit.getRange(`E${r}`).formulas = [[`=ROUND($C$21*(1+$C$22)^(A${r}-1),0)`]];
  unit.getRange(`F${r}`).formulas = [[`=ROUND(E${r}*$C$4*(1-$C$6),0)`]];
  unit.getRange(`G${r}`).formulas = [[`=ROUND(F${r}*$C$17-$C$18,0)`]];
  unit.getRange(`H${r}`).formulas = [[`=ROUND($D$21*(1+$D$22)^(A${r}-1),0)`]];
  unit.getRange(`I${r}`).formulas = [[`=ROUND(H${r}*$D$4*(1-$D$6),0)`]];
  unit.getRange(`J${r}`).formulas = [[`=ROUND(I${r}*$D$17-$D$18,0)`]];
}
unit.getRange("B23").formulas = [["=ROUND(SUM(C29:C40),0)"]];
unit.getRange("C23").formulas = [["=ROUND(SUM(F29:F40),0)"]];
unit.getRange("D23").formulas = [["=ROUND(SUM(I29:I40),0)"]];
unit.getRange("B24").formulas = [["=ROUND(SUM(D29:D40),0)"]];
unit.getRange("C24").formulas = [["=ROUND(SUM(G29:G40),0)"]];
unit.getRange("D24").formulas = [["=ROUND(SUM(J29:J40),0)"]];
unit.getRange("B3:D5").setNumberFormat("#,##0");
unit.getRange("B6:D6").setNumberFormat("0.0%");
unit.getRange("B7:D7").setNumberFormat("¥#,##0;[Red](¥#,##0);-");
unit.getRange("B8:D15").setNumberFormat("0.0%");
unit.getRange("B16:D16").setNumberFormat("¥#,##0;[Red](¥#,##0);-");
unit.getRange("B17:D17").setNumberFormat("0.0%");
unit.getRange("B18:D19").setNumberFormat("¥#,##0;[Red](¥#,##0);-");
unit.getRange("B20:D21").setNumberFormat("#,##0");
unit.getRange("B22:D22").setNumberFormat("0.0%");
unit.getRange("B23:D24").setNumberFormat("¥#,##0;[Red](¥#,##0);-");
unit.getRange("B29:B40").setNumberFormat("#,##0");
unit.getRange("E29:E40").setNumberFormat("#,##0");
unit.getRange("H29:H40").setNumberFormat("#,##0");
unit.getRange("C29:D40").setNumberFormat("¥#,##0;[Red](¥#,##0);-");
unit.getRange("F29:G40").setNumberFormat("¥#,##0;[Red](¥#,##0);-");
unit.getRange("I29:J40").setNumberFormat("¥#,##0;[Red](¥#,##0);-");
unit.freezePanes.freezeRows(2);
unit.getRange("A:A").format.columnWidth = 27;
unit.getRange("B:D").format.columnWidth = 18;
unit.getRange("E:J").format.columnWidth = 16;

// 90-day plan
titleBand(plan, "A1:H1", "90天验证计划 — 先验证，再成立重资产主体");
plan.getRange("A3:H3").values = [["阶段", "时间", "核心任务", "交付物", "预算CNY", "预算IDR", "继续门槛", "止损/决策"]];
styleHeader(plan.getRange("A3:H3"));
const planRows = [
  ["0 号阶段", "Day 1–30", "访谈买家、报价与合规扫描；确认进口商/经销商", "30个合格访谈；10份供应商报价；2个进口伙伴", 150000, null, "至少2条赛道各获5个明确采购痛点", "痛点弱或落地成本吞噬毛利，则砍掉该赛道"],
  ["1 号阶段", "Day 31–90", "三条赛道小规模并行试点；本地库存，不直邮", "MRO样品/备件；美妆3–5 SKU；摩托10–15 SKU", 450000, null, "至少一条赛道达到对应商业门槛", "总投入封顶45万；两月连续未达2项关键门槛即停"],
  ["2 号阶段", "Month 4–6", "只放大胜出赛道；招聘本地销售/售后", "可复制渠道＋单月贡献毛利转正", 1200000, null, "贡献毛利率达标且出现复购/渠道拉动", "没有复购或现金周期失控，不成立PT PMA"],
  ["3 号阶段", "Month 7–12", "补库存、团队和合规；必要时设立PT PMA", "1个主赛道＋1个期权赛道", 2200000, null, "滚动3月经营现金流接近盈亏平衡", "保留至少30%预算作退出和库存折价准备"],
];
plan.getRange("A4:H7").values = planRows;
plan.getRange("F4").formulas = [["=E4*'Sources'!$C$15"]];
plan.getRange("F4:F7").fillDown();
plan.getRange("A8:D8").merge();
plan.getRange("A8").values = [["合计（基准情景）"]];
plan.getRange("E8").formulas = [["=SUM(E4:E7)"]];
plan.getRange("F8").formulas = [["=SUM(F4:F7)"]];
plan.getRange("A8:H8").format = { fill: blue, font: { bold: true }, borders: { preset: "doubleBottom", style: "medium", color: navy } };
plan.getRange("E4:F8").format.numberFormat = "#,##0";
plan.getRange("F4:F8").format.font = { color: linkGreen };
plan.getRange("C4:D7").format.wrapText = true;
plan.getRange("G4:H7").format.wrapText = true;
plan.getRange("A4:H7").format.rowHeight = 46;
sectionBand(plan, "A11:H11", "赛道门槛（Day 90）");
plan.getRange("A12:H12").values = [["赛道", "付费验证", "毛利门槛", "获客/销售", "质量", "复购/渠道", "现金周期", "通过规则"]];
styleHeader(plan.getRange("A12:H12"));
plan.getRange("A13:H15").values = [
  ["工业MRO", "≥10个付费B2B客户", "落地毛利≥28%", "报价转订单≥20%", "重大售后事故=0", "60天内二次订单/扩品≥25%", "DSO≤45天", "8项中至少6项，且毛利/回款必须通过"],
  ["本地化美妆", "累计≥3,000单", "贡献毛利≥18%", "获客成本≤客单价25%", "退货≤5%；差评≤2%", "60天复购≥15%", "库存周转≤75天", "8项中至少6项，且合规/贡献毛利必须通过"],
  ["摩托后市场", "≥20个门店或1,500单", "贡献毛利≥15%", "广告+佣金≤客单价15%", "缺陷/退换≤2%", "门店补货率≥30%", "库存周转≤60天", "8项中至少6项，且质量/毛利必须通过"],
];
plan.getRange("A13:H15").format.wrapText = true;
plan.getRange("A13:H15").format.rowHeight = 42;
sectionBand(plan, "A18:H18", "预算原则");
plan.getRange("A19:H21").merge(true);
plan.getRange("A19").values = [["1. 400万元基准预算不是一次性承诺；前90天只授权60万元。"]];
plan.getRange("A20").values = [["2. PT PMA实缴资本约IDR25亿，按当前汇率约95万元人民币；它不是全部费用，但会显著占用启动现金。"]];
plan.getRange("A21").values = [["3. 先用合规进口商/经销商验证，胜出后再决定PT PMA或本地OEM；避免以名义股东规避外资规定。"]];
plan.getRange("A19:H21").format = { fill: paleYellow, wrapText: true, font: { color: dark } };
plan.getRange("A19:H21").format.rowHeight = 28;
plan.freezePanes.freezeRows(3);
plan.getRange("A:A").format.columnWidth = 14;
plan.getRange("B:B").format.columnWidth = 15;
plan.getRange("C:D").format.columnWidth = 34;
plan.getRange("E:F").format.columnWidth = 16;
plan.getRange("G:H").format.columnWidth = 36;

// Regulatory
titleBand(regulatory, "A1:G1", "进入印尼前的监管与结构检查");
regulatory.getRange("A3:G3").values = [["事项", "适用赛道", "当前规则/风险", "进入前动作", "最晚完成", "红旗", "来源ID"]];
styleHeader(regulatory.getRange("A3:G3"));
regulatory.getRange("A4:G10").values = [
  ["低价跨境电商", "全部低价消费品", "境外卖家经跨境PMSE直接向印尼销售外国成品，平台原则上须执行FOB USD100/件最低价", "采用批量进口＋印尼仓；核对例外清单", "上线前", "继续依赖中国逐单直邮", "S09 / Permendag 19/2026"],
  ["清真认证", "食品、美妆、健康、部分家清/日用品", "PP 39/2021分阶段期至2026-10-17；应把2026-10-18起的强制要求纳入计划", "确认分类、认证主体和境外机构互认路径", "SKU冻结前", "配方/原料无法追溯", "PP 39/2021"],
  ["BPOM", "美妆、食品、保健品", "进口成品需当地合格主体办理产品通知/注册及进口文件；功效宣称受约束", "逐SKU做配方、标签、宣称和文件差距分析", "首批生产前", "先生产后备案；夸大功效", "BPOM"],
  ["SNI/技术标准", "头盔、电气件、特定设备", "部分产品适用强制SNI或技术要求", "按HS/KBLI和产品型号核验", "报价/下单前", "供应商只提供中国证书", "Kemendag/BSN"],
  ["进口主体与NIB/API", "全部进口业务", "需由具备相应许可和海关能力的印尼主体进口", "尽调进口商：NIB、API、税务、历史清关", "签合同前", "借牌进口、报价不含税费", "OSS/BKPM"],
  ["PT PMA资本", "自营团队/长期经营", "最低实缴资本一般为IDR25亿；投资计划通常> IDR100亿/5位KBLI/项目地点", "验证后再设立；由印尼律师按KBLI确认", "扩张决策前", "把投资计划误认为全部需一次缴清；忽视资本占用", "S12"],
  ["赊销与回款", "工业MRO/设备", "B2B回款期可能侵蚀利润和现金", "设信用额度、预付款和逾期停供条款", "首单前", "无担保长账期；单一大客户", "内部风控"],
];
regulatory.getRange("A4:G10").format.wrapText = true;
regulatory.getRange("A4:G10").format.verticalAlignment = "top";
regulatory.getRange("A4:G10").format.rowHeight = 58;
regulatory.getRange("F4:F10").format.fill = paleRed;
regulatory.freezePanes.freezeRows(3);
regulatory.getRange("A:B").format.columnWidth = 24;
regulatory.getRange("C:D").format.columnWidth = 42;
regulatory.getRange("E:E").format.columnWidth = 18;
regulatory.getRange("F:F").format.columnWidth = 35;
regulatory.getRange("G:G").format.columnWidth = 24;

// Summary
titleBand(summary, "A1:J1", "印尼高增长行业机会 — 决策版");
summary.getRange("A2:J2").merge();
summary.getRange("A2").values = [["基准对象：中国供应链/品牌团队｜首阶段预算 200–500 万元｜先验证、后设立重资产主体｜更新至 2026-08-30"]];
summary.getRange("A2").format = { font: { color: "#666666", italic: true }, wrapText: true };

summary.getRange("A4:B4").merge();
summary.getRange("A4").values = [["首选赛道"]];
summary.getRange("A5:B6").merge();
summary.getRange("A5").formulas = [["='Scorecard'!B5"]];
summary.getRange("A5").format = { fill: paleGreen, font: { bold: true, color: navy, size: 16 }, verticalAlignment: "center", horizontalAlignment: "center", wrapText: true };

summary.getRange("C4:D4").merge();
summary.getRange("C4").values = [["综合评分"]];
summary.getRange("C5:D6").merge();
summary.getRange("C5").formulas = [["='Scorecard'!K5"]];
summary.getRange("C5").format = { fill: paleGreen, font: { bold: true, color: navy, size: 18 }, verticalAlignment: "center", horizontalAlignment: "center" };
summary.getRange("C5").format.numberFormat = "0.00";

summary.getRange("E4:F4").merge();
summary.getRange("E4").values = [["前90天授权预算"]];
summary.getRange("E5:F6").merge();
summary.getRange("E5").formulas = [["=\"¥\"&TEXT(SUM('90-Day Plan'!E4:E5),\"#,##0\")"]];
summary.getRange("E5").format = { fill: paleYellow, font: { bold: true, color: navy, size: 16 }, verticalAlignment: "center", horizontalAlignment: "center" };
summary.getRange("E5:F6").setNumberFormat("¥#,##0");

summary.getRange("G4:H4").merge();
summary.getRange("G4").values = [["PT PMA实缴资本约"]];
summary.getRange("G5:H6").merge();
summary.getRange("G5").formulas = [["=\"¥\"&TEXT(ROUND('Sources'!C16*1000000/'Sources'!C15,0),\"#,##0\")"]];
summary.getRange("G5").format = { fill: paleBlue, font: { bold: true, color: navy, size: 16 }, verticalAlignment: "center", horizontalAlignment: "center" };
summary.getRange("G5:H6").setNumberFormat("¥#,##0");

summary.getRange("I4:J4").merge();
summary.getRange("I4").values = [["模型状态"]];
summary.getRange("I5:J6").merge();
summary.getRange("I5").formulas = [["='Checks'!F3"]];
summary.getRange("I5").format = { fill: paleGreen, font: { bold: true, color: "#006100", size: 16 }, verticalAlignment: "center", horizontalAlignment: "center" };

sectionBand(summary, "A9:J9", "核心判断");
summary.getRange("A10:J10").merge();
summary.getRange("A11:J11").merge();
summary.getRange("A12:J12").merge();
summary.getRange("A10").values = [["1. 最优不是“增速最高的消费单品”，而是工业MRO/自动化：需求由下游投资和资本品进口驱动，中国供应链优势强，且可用库存与售后建立壁垒。"]];
summary.getRange("A11").values = [["2. 消费品里，美妆规模最大、家清增速更快，但视频电商客单仅约USD4.5–6，必须使用组合装、本地库存/生产和达人分销来保护单位经济。"]];
summary.getRange("A12").values = [["3. 400万元预算应采用阶段授权：前90天仅60万元；PT PMA约95万元实缴资本会占用现金，因此应在胜出赛道出现复购和贡献毛利后再设立。"]];
summary.getRange("A10:J12").format = { fill: "#F8FAFC", wrapText: true, font: { color: dark } };
summary.getRange("A10:J12").format.rowHeight = 32;

sectionBand(summary, "A15:F15", "优先级与建议切口");
summary.getRange("A16:F16").values = [["排名", "赛道", "综合分", "需求证据", "建议切口", "建议"]];
styleHeader(summary.getRange("A16:F16"));
for (let i = 0; i < 6; i++) {
  const r = 17 + i;
  const sr = 5 + i;
  summary.getRange(`A${r}`).formulas = [[`='Scorecard'!A${sr}`]];
  summary.getRange(`B${r}`).formulas = [[`='Scorecard'!B${sr}`]];
  summary.getRange(`C${r}`).formulas = [[`='Scorecard'!K${sr}`]];
  summary.getRange(`D${r}`).formulas = [[`='Scorecard'!L${sr}`]];
  summary.getRange(`E${r}`).formulas = [[`='Scorecard'!M${sr}`]];
  summary.getRange(`F${r}`).values = [[i < 2 ? "立即验证" : i < 4 ? "作为备选" : "暂缓重投入"]];
}
summary.getRange("C17:C22").format.numberFormat = "0.00";
summary.getRange("A17:E22").format.font = { color: linkGreen };
summary.getRange("D17:F22").format.wrapText = true;
summary.getRange("F17:F18").format.fill = paleGreen;
summary.getRange("F19:F20").format.fill = paleYellow;
summary.getRange("F21:F22").format.fill = grey;

sectionBand(summary, "H15:J15", "宏观与渠道信号");
summary.getRange("H16:I16").values = [["指标", "值"]];
styleHeader(summary.getRange("H16:I16"));
summary.getRange("H17:H22").values = [["2026 Q2 GDP同比"], ["2026 H1住宿餐饮"], ["2025下游投资同比"], ["中国占进口"], ["H1 2026电商FMCG"], ["视频电商交易量同比"]];
summary.getRange("I17").formulas = [["='Sources'!C5/100"]];
summary.getRange("I18").formulas = [["='Sources'!C6/100"]];
summary.getRange("I19").formulas = [["='Sources'!C10/100"]];
summary.getRange("I20").formulas = [["='Sources'!C8/100"]];
summary.getRange("I21").formulas = [["='Sources'!C11"]];
summary.getRange("I22").formulas = [["='Sources'!C13/100"]];
summary.getRange("I17:I20").format.numberFormat = "0.0%";
summary.getRange("I21").format.numberFormat = "0.0 \"IDR tn\"";
summary.getRange("I22").format.numberFormat = "0%";
summary.getRange("I17:I22").format.font = { color: linkGreen, bold: true };

summary.getRange("H25:I25").values = [["赛道", "综合分"]];
styleHeader(summary.getRange("H25:I25"));
for (let i = 0; i < 6; i++) {
  const r = 26 + i;
  const sr = 5 + i;
  summary.getRange(`H${r}:I${r}`).formulas = [[`='Scorecard'!B${sr}`, `='Scorecard'!K${sr}`]];
}
summary.getRange("I26:I31").format.numberFormat = "0.00";
const rankChart = summary.charts.add("bar", summary.getRange("H25:I31"));
rankChart.title = "机会综合评分（10分制）";
rankChart.hasLegend = false;
rankChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
rankChart.yAxis = { numberFormatCode: "0.0", min: 0, max: 10 };
rankChart.setPosition("H33", "J47");

summary.freezePanes.freezeRows(2);
summary.getRange("A:A").format.columnWidth = 9;
summary.getRange("B:B").format.columnWidth = 24;
summary.getRange("C:C").format.columnWidth = 12;
summary.getRange("D:D").format.columnWidth = 31;
summary.getRange("E:E").format.columnWidth = 42;
summary.getRange("F:F").format.columnWidth = 15;
summary.getRange("G:G").format.columnWidth = 5;
summary.getRange("H:H").format.columnWidth = 28;
summary.getRange("I:I").format.columnWidth = 18;
summary.getRange("J:J").format.columnWidth = 12;

// Checks
titleBand(checks, "A1:G1", "模型检查");
checks.getRange("A3:E3").merge();
checks.getRange("A3").values = [["总体状态"]];
checks.getRange("F3:G3").merge();
checks.getRange("F3").formulas = [["=IF(AND(F6=\"OK\",F7=\"OK\",F8=\"OK\",F9=\"OK\",F10=\"OK\"),\"OK\",\"REVIEW\")"]];
checks.getRange("A3:G3").format = { fill: paleGreen, font: { bold: true, size: 14 }, borders: { preset: "outside", style: "medium", color: teal } };
checks.getRange("A5:G5").values = [["检查", "实际", "预期", "差异", "容差", "状态", "说明"]];
styleHeader(checks.getRange("A5:G5"));
checks.getRange("A6:A10").values = [["权重合计"], ["评分范围"], ["预算合计"], ["三赛道贡献毛利"], ["来源完整性"]];
checks.getRange("B6").formulas = [["=SUM('Scorecard'!C3:J3)"]];
checks.getRange("C6:E6").values = [[1, 0, 0.0001]];
checks.getRange("F6").formulas = [["=IF(ABS(D6)<=E6,\"OK\",\"REVIEW\")"]];
checks.getRange("G6").values = [["权重应等于100%"]];
checks.getRange("D6").formulas = [["=B6-C6"]];
checks.getRange("B7").formulas = [["=MIN('Scorecard'!C5:J10)"]];
checks.getRange("C7:E7").values = [[1, null, 0]];
checks.getRange("D7").formulas = [["=MAX('Scorecard'!C5:J10)"]];
checks.getRange("F7").formulas = [["=IF(AND(B7>=1,D7<=10),\"OK\",\"REVIEW\")"]];
checks.getRange("G7").values = [["所有评分须在1–10之间"]];
checks.getRange("B8").formulas = [["='90-Day Plan'!E8"]];
checks.getRange("C8:E8").values = [[4000000, null, 1]];
checks.getRange("D8").formulas = [["=B8-C8"]];
checks.getRange("F8").formulas = [["=IF(ABS(D8)<=E8,\"OK\",\"REVIEW\")"]];
checks.getRange("G8").values = [["基准预算应为400万元"]];
checks.getRange("B9").formulas = [["=MIN('Unit Economics'!B17:D17)"]];
checks.getRange("C9:E9").values = [[0, null, 0]];
checks.getRange("D9").formulas = [["=B9-C9"]];
checks.getRange("F9").formulas = [["=IF(B9>0,\"OK\",\"REVIEW\")"]];
checks.getRange("G9").values = [["三赛道贡献毛利率均须为正"]];
checks.getRange("B10").formulas = [["=COUNTA('Sources'!G5:G17)"]];
checks.getRange("C10:E10").values = [[13, null, 0]];
checks.getRange("D10").formulas = [["=B10-C10"]];
checks.getRange("F10").formulas = [["=IF(B10=C10,\"OK\",\"REVIEW\")"]];
checks.getRange("G10").values = [["13条外部输入均应有URL"]];
checks.getRange("B6:E6").format.numberFormat = "0.0%";
checks.getRange("B8:D8").format.numberFormat = "¥#,##0";
checks.getRange("B9:D9").format.numberFormat = "0.0%";
checks.getRange("F6:F10").conditionalFormats.add("containsText", { text: "OK", format: { fill: "#C6EFCE", font: { color: "#006100", bold: true } } });
checks.getRange("F6:F10").conditionalFormats.add("containsText", { text: "REVIEW", format: { fill: "#FFC7CE", font: { color: "#9C0006", bold: true } } });
checks.getRange("A6:G10").format.wrapText = true;
checks.getRange("A:A").format.columnWidth = 24;
checks.getRange("B:F").format.columnWidth = 16;
checks.getRange("G:G").format.columnWidth = 42;

// Cross-sheet/formula styling
summary.getRange("A5:I31").format.verticalAlignment = "center";
summary.getRange("A16:F22").format.borders = { preset: "inside", style: "thin", color: "#D9E2F3" };
unit.getRange("A3:A24").format.font = { color: dark };
unit.getRange("B5:D24").format.font = { color: "#000000" };
unit.getRange("B3:D4").format.font = { color: inputBlue };
unit.getRange("B6:D6").format.font = { color: inputBlue };
unit.getRange("B8:D14").format.font = { color: inputBlue };
unit.getRange("B18:D18").format.font = { color: inputBlue };
unit.getRange("B21:D22").format.font = { color: inputBlue };

// Export and verification outputs
await fs.mkdir(outputDir, { recursive: true });
const inspectSummary = await wb.inspect({ kind: "table", range: "Summary!A1:J31", include: "values,formulas", tableMaxRows: 31, tableMaxCols: 10, maxChars: 12000 });
console.log("SUMMARY_INSPECT\n" + inspectSummary.ndjson);
const inspectUnit = await wb.inspect({ kind: "table", range: "Unit Economics!A1:J40", include: "values,formulas", tableMaxRows: 40, tableMaxCols: 10, maxChars: 12000 });
console.log("UNIT_INSPECT\n" + inspectUnit.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" });
console.log("ERROR_SCAN\n" + errors.ndjson);

for (const [sheetName, fileName, range] of [
  ["Summary", "preview_summary.png", "A1:J47"],
  ["Scorecard", "preview_scorecard.png", "A1:M10"],
  ["Unit Economics", "preview_unit.png", "A1:J40"],
  ["90-Day Plan", "preview_plan.png", "A1:H21"],
  ["Regulatory", "preview_regulatory.png", "A1:G10"],
  ["Sources", "preview_sources.png", "A1:H17"],
  ["Checks", "preview_checks.png", "A1:G10"],
]) {
  const blob = await wb.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${fileName}`, new Uint8Array(await blob.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(`SAVED ${outputPath}`);
