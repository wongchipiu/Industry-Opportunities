import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const outputDir = path.join(repoRoot, "outputs/019fdcb6-ce96-70e1-951a-011c3a45e8ec");
const outputPath = `${outputDir}/医药行业基础架构与新技术赛道量化分析_2026-08-07.xlsx`;
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const overview = workbook.worksheets.add("总览");
const architecture = workbook.worksheets.add("产业架构");
const tracks = workbook.worksheets.add("赛道评分");
const players = workbook.worksheets.add("玩家矩阵");
const shovels = workbook.worksheets.add("卖铲子映射");
const companies = workbook.worksheets.add("公司量化");
const model = workbook.worksheets.add("评分模型");
const sources = workbook.worksheets.add("来源与口径");
const checks = workbook.worksheets.add("检查");

const COLORS = {
  navy: "#16324F",
  blue: "#2563EB",
  teal: "#0F766E",
  green: "#15803D",
  lightGreen: "#DCFCE7",
  amber: "#D97706",
  lightAmber: "#FEF3C7",
  red: "#B91C1C",
  lightRed: "#FEE2E2",
  slate: "#475569",
  lightSlate: "#E2E8F0",
  pale: "#F8FAFC",
  white: "#FFFFFF",
  black: "#111827",
  inputBlue: "#0000FF",
  linkGreen: "#008000",
};

function title(sheet, range, text) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 16 },
    verticalAlignment: "center",
  };
  sheet.getRange(range).format.rowHeight = 32;
}

function section(sheet, range, text) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 11 },
    verticalAlignment: "center",
  };
  sheet.getRange(range).format.rowHeight = 24;
}

function header(sheet, range) {
  sheet.getRange(range).format = {
    fill: COLORS.slate,
    font: { bold: true, color: COLORS.white },
    verticalAlignment: "center",
    horizontalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#CBD5E1" },
  };
  sheet.getRange(range).format.rowHeight = 30;
}

function body(sheet, range) {
  sheet.getRange(range).format = {
    font: { color: COLORS.black, size: 10 },
    verticalAlignment: "top",
    wrapText: true,
    borders: { insideHorizontal: { style: "thin", color: "#E5E7EB" } },
  };
}

function setWidths(sheet, specs) {
  for (const [col, width] of Object.entries(specs)) sheet.getRange(`${col}:${col}`).format.columnWidth = width;
}

for (const sheet of [overview, architecture, tracks, players, shovels, companies, model, sources, checks]) {
  sheet.showGridLines = false;
}

// -------------------- 评分模型 --------------------
title(model, "A1:D1", "评分模型与可编辑权重");
model.getRange("A2:D2").values = [["说明", "蓝字为可调整输入；所有综合分均由公式计算", "数据日期", new Date("2026-08-07T00:00:00+08:00")]];
model.getRange("D2").format.numberFormat = "yyyy-mm-dd";

section(model, "A3:D3", "创新药赛道吸引力权重");
model.getRange("A4:C8").values = [
  ["维度", "权重", "定义"],
  ["商业化验证", 0.25, "获批产品规模与收入深度"],
  ["管线广度", 0.20, "活跃干预性临床试验数，对数标准化"],
  ["增长动量", 0.20, "头部产品/PLATFORM 2025 年收入趋势"],
  ["临床/支付韧性", 0.20, "安全性、可及性、报销与竞争格局"],
];
model.getRange("A9:C9").values = [["上市公司可投性", 0.15, "A股/美股可用标的数量与质量"]];
header(model, "A4:C4");
body(model, "A5:C9");
model.getRange("B5:B9").format = { font: { color: COLORS.inputBlue }, fill: "#EFF6FF", numberFormat: "0%" };

section(model, "A11:D11", "卖铲子吸引力权重");
model.getRange("A12:C17").values = [
  ["维度", "权重", "定义"],
  ["工艺复杂度", 0.25, "工艺开发、放大与分析难度"],
  ["耗材复购", 0.20, "树脂、培养基、试剂、包材等重复收入"],
  ["商业需求可见度", 0.25, "终端放量向上游传导的确定性"],
  ["产能/价格纪律", 0.15, "产能利用率、供给过剩与价格竞争"],
  ["跨项目复用", 0.15, "平台是否可服务多客户/多分子"],
];
header(model, "A12:C12");
body(model, "A13:C17");
model.getRange("B13:B17").format = { font: { color: COLORS.inputBlue }, fill: "#EFF6FF", numberFormat: "0%" };

section(model, "A19:D19", "组合权重");
model.getRange("A20:C22").values = [
  ["组件", "权重", "用途"],
  ["创新药赛道分", 0.55, "衡量终端技术胜率与收入空间"],
  ["卖铲子分", 0.45, "衡量工具、耗材与外包环节的确定性"],
];
header(model, "A20:C20");
body(model, "A21:C22");
model.getRange("B21:B22").format = { font: { color: COLORS.inputBlue }, fill: "#EFF6FF", numberFormat: "0%" };

section(model, "A24:D24", "公司量化权重");
model.getRange("A25:C30").values = [
  ["维度", "权重", "定义"],
  ["收入动量", 0.30, "FY2025 收入增长分档"],
  ["盈利质量", 0.20, "GAAP/经调整利润与现金流"],
  ["壁垒/关键性", 0.25, "客户验证、监管、IP、转换成本"],
  ["跨赛道分散", 0.15, "多客户、多模态、多地域覆盖"],
  ["估值舒适度", 0.10, "当前价格 / EPS 的粗略 P/E 分档"],
];
header(model, "A25:C25");
body(model, "A26:C30");
model.getRange("B26:B30").format = { font: { color: COLORS.inputBlue }, fill: "#EFF6FF", numberFormat: "0%" };

section(model, "A32:D32", "收入增长分档");
model.getRange("A33:C37").values = [
  ["增长下限", "得分", "解释"],
  [0.20, 5, ">=20%"],
  [0.12, 4, "12%-20%"],
  [0.05, 3, "5%-12%"],
  [0.00, 2, "0%-5%"],
];
model.getRange("A38:C38").values = [[-9.99, 1, "<0%"]];
header(model, "A33:C33");
body(model, "A34:C38");
model.getRange("A34:B38").format.font = { color: COLORS.inputBlue };
model.getRange("A34:A38").format.numberFormat = "0.0%";

section(model, "A40:D40", "P/E 估值分档");
model.getRange("A41:C45").values = [
  ["P/E 上限", "得分", "解释"],
  [20, 5, "<=20x"],
  [30, 4, "20x-30x"],
  [40, 3, "30x-40x"],
  [60, 2, "40x-60x"],
];
model.getRange("A46:C46").values = [[9999, 1, ">60x，或亏损/无法计算"]];
header(model, "A41:C41");
body(model, "A42:C46");
model.getRange("A42:B46").format.font = { color: COLORS.inputBlue };
setWidths(model, { A: 24, B: 14, C: 58, D: 16 });
model.freezePanes.freezeRows(3);

// -------------------- 赛道评分 --------------------
title(tracks, "A1:R1", "GLP-1 / ADC / 双抗 / CGT / mRNA 赛道量化评分");
tracks.getRange("A2:R2").merge();
tracks.getRange("A2:R2").values = [["临床试验数为 ClinicalTrials.gov 2026-08-07 活跃干预性试验，仅作热度代理；1-5 分中 5 为最优。蓝字为可编辑判断。"]];
tracks.getRange("A2:R2").format = { fill: COLORS.pale, font: { color: COLORS.slate, italic: true }, wrapText: true };
tracks.getRange("A4:R4").values = [[
  "赛道", "活跃干预性试验", "FY2025 商业信号", "商业验证", "管线广度", "增长动量", "临床/支付韧性", "上市公司可投性", "创新药分", "工艺复杂度", "耗材复购", "需求可见度", "产能/价格纪律", "跨项目复用", "卖铲子分", "综合分", "核心结论", "来源ID"
]];
header(tracks, "A4:R4");
tracks.getRange("A5:R9").values = [
  ["GLP-1", 777, "Lilly Mounjaro+Zepbound $36.5bn；Novo 三大司美格鲁肽品牌 DKK229.5bn", 5, null, 5, 4, 5, null, 4, 4, 5, 4, 3, null, null, "近端收入确定性最强，胶原铲子在多肽API/CDMO、纯化与注射器件", "S02,S07,S08"],
  ["ADC", 565, "Enhertu 9M25 合并销售 $3.575bn；Trodelvy FY25 $1.397bn", 4, null, 4, 3, 4, null, 5, 5, 4, 4, 4, null, null, "单药风险高，但 linker/payload/偶联/分析的卖铲子弹性最强", "S03,S09,S10"],
  ["双抗", 485, "Roche Hemlibra+Vabysmo CHF8.856bn；Amgen Blincyto $1.559bn", 4, null, 4, 4, 4, null, 3, 5, 4, 3, 5, null, null, "商业验证已越过肿瘤，生产多可复用成熟大分子平台", "S04,S11,S12"],
  ["CGT", 2061, "Carvykti $1.887bn(+96%)；Gilead 细胞治疗 $1.839bn(-7%)", 3, null, 3, 2, 3, null, 5, 4, 3, 2, 4, null, null, "管线最宽但商业化分化，需关注产能利用率、收治中心与安全性", "S05,S10,S13,S14"],
  ["mRNA", 95, "Moderna FY25 收入 $1.9bn(-40%)；Maravai TriLink $119.8m(-39%)", 2, null, 2, 2, 2, null, 4, 4, 2, 1, 3, null, null, "平台仍有期权价值，但大流行后产能过剩和需求收缩尚未出清", "S06,S15,S16,S27"],
];
body(tracks, "A5:R9");
tracks.getRange("D5:D9").format.font = { color: COLORS.inputBlue };
tracks.getRange("F5:H9").format.font = { color: COLORS.inputBlue };
tracks.getRange("J5:N9").format.font = { color: COLORS.inputBlue };
tracks.getRange("D5:H9").format.fill = "#EFF6FF";
tracks.getRange("J5:N9").format.fill = "#EFF6FF";
for (let r = 5; r <= 9; r++) {
  tracks.getRange(`E${r}`).formulas = [[`=1+4*(LN(B${r})-LN(MIN($B$5:$B$9)))/(LN(MAX($B$5:$B$9))-LN(MIN($B$5:$B$9)))`]];
  tracks.getRange(`I${r}`).formulas = [[`=20*SUMPRODUCT(D${r}:H${r},'评分模型'!$B$5:$B$9)`]];
  tracks.getRange(`O${r}`).formulas = [[`=20*SUMPRODUCT(J${r}:N${r},'评分模型'!$B$13:$B$17)`]];
  tracks.getRange(`P${r}`).formulas = [[`=I${r}*'评分模型'!$B$21+O${r}*'评分模型'!$B$22`]];
}
tracks.getRange("E5:E9").format.numberFormat = "0.0";
tracks.getRange("I5:I9").format.numberFormat = "0.0";
tracks.getRange("O5:P9").format.numberFormat = "0.0";
tracks.getRange("D5:H9").format.horizontalAlignment = "center";
tracks.getRange("I5:P9").format.horizontalAlignment = "center";
tracks.getRange("I5:I9").conditionalFormats.add("colorScale", { colors: ["#FEE2E2", "#FEF3C7", "#DCFCE7"] });
tracks.getRange("O5:P9").conditionalFormats.add("colorScale", { colors: ["#FEE2E2", "#FEF3C7", "#DCFCE7"] });
setWidths(tracks, { A: 12, B: 14, C: 48, D: 11, E: 11, F: 11, G: 14, H: 13, I: 11, J: 11, K: 11, L: 12, M: 14, N: 12, O: 11, P: 11, Q: 50, R: 17 });
tracks.freezePanes.freezeRows(4);
tracks.freezePanes.freezeColumns(1);

// -------------------- 产业架构 --------------------
title(architecture, "A1:H1", "医药行业的基础架构：资产、服务、制造与工具");
section(architecture, "A3:H3", "从靶点到患者的价值链");
architecture.getRange("A4:H4").values = [["层级", "主要角色", "收入模式", "核心壁垒", "关键KPI", "轻/重资产", "主要风险", "代表公司"]];
header(architecture, "A4:H4");
architecture.getRange("A5:H14").values = [
  ["资金/资产持有者", "Biotech/Pharma", "产品收入、里程碑、授权金", "IP、临床数据、商业化", "管线阶段、PoS、销售峰值", "前期轻/商业化变重", "临床失败、专利、支付", "Lilly、Novo、恒瑞、百济"],
  ["发现赋能", "药物发现CRO/AI/化合物库", "FTE/FFS、项目包、里程碑", "人才、数据、通量与交付", "人均产出、利用率、R-to-D转化", "轻-中", "价格竞争、biotech融资", "药明康德、康龙化成"],
  ["临床前", "DMPK/毒理/模型CRO", "项目计价+容量使用", "GLP资质、模型、法规经验", "订单、产能利用率、报告周期", "中-重", "实验动物供应、产能过剩", "Charles River、昭衍新药"],
  ["临床开发", "临床CRO/SMO/数据", "FFS、全包项目、程序化收入", "全球网络、中心/患者招募、数据", "Backlog、Book-to-bill、取消率", "轻-中", "试验延迟/取消、人员流动", "IQVIA、Medpace、泰格"],
  ["CMO", "按客户成熟工艺生产", "批次×单价", "GMP、质量、放大与交付", "利用率、批次释放、偏差率", "重", "利用率、客户集中、质量事故", "大型商业化生产平台"],
  ["CDMO", "工艺开发+生产", "开发费+临床批+商业批", "工艺IP、法规、转移成本", "项目漏斗、后期/商业化占比", "重", "技术转移、客户产品失败", "Thermo Fisher、凯莱英"],
  ["CRDMO", "发现→研究→开发→生产", "多阶段交叉销售", "跟随分子、平台协同、一体化", "新增分子、转化率、在手订单", "中-重", "地缘、平台复杂度、执行", "药明康德、康龙化成"],
  ["科研工具/试剂", "仪器、耗材、试剂、软件", "装机+耗材复购+服务", "安装基数、客户验证、应用know-how", "有机增长、耗材渗透、毛利率", "中", "去库存、科研资助、产品替代", "Danaher、Repligen、纳微"],
  ["包材/给药", "胶塞、预充针、自注射器", "高价值组件+合同制造", "客户验证、药械组合、规模", "HVP占比、有机增长", "中-重", "客户集中、产线爬坡", "West Pharma、Stevanato"],
  ["商业化/支付", "CSO、渠道、PBM/保险", "服务费、差价、佣金", "渠道、付款方谈判、现实世界数据", "准入、处方量、患者覆盖", "轻", "政策、控费、合规", "IQVIA等数据/商业化服务"],
];
body(architecture, "A5:H14");

section(architecture, "A16:H16", "CRO / CMO / CDMO / CRDMO 经济性差异");
architecture.getRange("A17:H17").values = [["模式", "客户买什么", "粗略收入式", "锁定性", "定价权", "资本开支", "最应关注KPI", "投资含义"]];
header(architecture, "A17:H17");
architecture.getRange("A18:H23").values = [
  ["CRO", "人员、时间、数据与法规执行", "FTE×利用率×单价 + pass-through", "中", "中", "低-中", "Backlog、book-to-bill、取消率", "更像专业服务；收入滞后于融资周期"],
  ["SMO", "临床现场执行与CRC", "中心数×服务费", "低-中", "低", "低", "中心启动、人均管理中心数", "劳动密集，规模不等于壁垒"],
  ["CMO", "已定型的合规产能", "批次×单价", "中", "低-中", "高", "利用率、良率、客户集中", "典型制造周期，产能纪律决定回报"],
  ["CDMO", "工艺与产能的一体化", "开发费+临床批+商业批", "高", "中-高", "高", "后期项目数、商业化收入", "较CMO更有黏性，但仍受利用率影响"],
  ["CRDMO", "端到端平台与分子漏斗", "多阶段收入+交叉销售", "很高", "高", "中-高", "R-to-D转化、在手订单、每客户收入", "平台最优，但要防地缘与组织复杂度"],
  ["工具/耗材", "经验证的工艺节点", "装机+一次性耗材×使用量", "高", "中-高", "中", "耗材复购、有机增长、毛利率", "最接近“卖铲子”，但要辨别周期库存"],
];
body(architecture, "A18:H23");
setWidths(architecture, { A: 20, B: 31, C: 34, D: 16, E: 16, F: 15, G: 30, H: 38 });
architecture.freezePanes.freezeRows(4);

// -------------------- 玩家矩阵 --------------------
title(players, "A1:J1", "头部玩家与可投暴露矩阵（A股、美股，港股/其他作必要补充）");
players.getRange("A3:J3").values = [["赛道", "角色", "市场", "公司", "代码", "核心产品/能力", "头部依据", "卖铲子相关性", "主要风险", "来源ID"]];
header(players, "A3:J3");
const playerRows = [
  ["GLP-1", "药品", "美股", "Eli Lilly", "LLY", "Mounjaro/Zepbound；retatrutide/orforglipron", "2025年两品牌合计$36.5bn", "低（终端需求源）", "价格、产能、竞品", "S07"],
  ["GLP-1", "药品", "美股ADR", "Novo Nordisk", "NVO", "Ozempic/Wegovy/Rybelsus", "2025三品牌合计DKK229.5bn", "低", "竞争加剧、增速放缓", "S08"],
  ["GLP-1", "药品", "美股", "Amgen", "AMGN", "MariTide", "长效肥胖管线", "低", "临床与安全性", "S12"],
  ["GLP-1", "药品", "美股", "Viking Therapeutics", "VKTX", "VK2735", "临床阶段纯肥胖暴露", "低", "单管线、融资", "S39"],
  ["GLP-1", "药品", "A股", "恒瑞医药", "600276.SH", "HRS9531等代谢管线", "国内综合创新药平台", "低", "临床/竞争、定价", "S40"],
  ["GLP-1", "药品", "A股", "华东医药", "000963.SZ", "利拉鲁肽及多靶点代谢布局", "国内商业化与管线并行", "中（制剂/产能）", "同质化与市场准入", "S41"],
  ["GLP-1", "药品", "A股", "博瑞医药", "688166.SH", "BGM0504", "国内多靶点候选者", "低", "临床/融资", "S42"],
  ["GLP-1", "API/CDMO", "A股", "药明康德", "603259.SH", "TIDES多肽/寡核苷酸", "2025 TIDES收入人民币113.7亿，+96%", "很高", "地缘政策、扩产执行", "S17"],
  ["GLP-1", "API/CDMO", "A股", "ST诺泰", "688076.SH", "多肽原料药/CDMO", "2025收入19.43亿，+19.6%", "很高", "ST/治理风险、利润下滑", "S32"],
  ["GLP-1", "API/CDMO", "A股", "凯莱英", "002821.SZ", "化学大分子/多肽工艺", "CDMO平台向新分子扩展", "高", "产能爬坡、客户项目波动", "S20"],
  ["ADC", "药品", "美股", "AstraZeneca / Daiichi Sankyo", "AZN / 4568.T", "Enhertu/Datroway", "Enhertu 9M25合并销售$3.575bn", "中", "ILD/毒性、治疗窗、竞争", "S10"],
  ["ADC", "药品", "美股", "Pfizer / Seagen", "PFE", "Padcev、Adcetris等", "并购Seagen后的商业化ADC组合", "中", "并购回报、产品集中", "S36"],
  ["ADC", "药品", "美股", "Gilead", "GILD", "Trodelvy", "FY25收入$1.397bn，+6%", "低", "适应症扩展与竞争", "S09"],
  ["ADC", "药品", "美股", "AbbVie", "ABBV", "Elahere及后续ADC管线", "并购ImmunoGen获得商业化平台", "中", "集成与管线成功率", "S37"],
  ["ADC", "药品", "A股", "恒瑞医药", "600276.SH", "多个ADC临床管线", "国内综合平台+出海能力", "中", "管线取舍与出海执行", "S40"],
  ["ADC", "药品", "A股", "科伦药业", "002422.SZ", "SKB264等ADC", "科伦博泰平台授权与临床进展", "中", "临床、授权依赖", "S43"],
  ["ADC", "药品", "A股", "百利天恒", "688506.SH", "iza-bren/BL-B01D1等", "双抗ADC技术与大额授权", "中-高", "单管线权重大、临床风险", "S44"],
  ["ADC", "药品", "A股/H股", "荣昌生物", "688331.SH", "Disitamab vedotin", "国内商业化ADC先行者", "中", "现金流与商业化效率", "S45"],
  ["ADC", "CDMO", "港股", "药明合联", "2268.HK", "偶联药CRDMO", "专注ADC/其他偶联药的端到端平台", "很高", "估值、客户/政策", "S46"],
  ["双抗", "药品", "其他", "Roche", "ROG.SW", "Hemlibra/Vabysmo、Lunsumio/Columvi", "FY25 Hemlibra CHF4.754bn; Vabysmo CHF4.102bn", "中", "不同机制间风险差异大", "S11"],
  ["双抗", "药品", "美股", "Amgen", "AMGN", "Blincyto/Imdelltra", "FY25 Blincyto $1.559bn，+28%", "中", "CRS、给药便利性", "S12"],
  ["双抗", "药品", "美股", "Johnson & Johnson", "JNJ", "Tecvayli/Talvey/Rybrevant", "FY25三品种分别$670m/$463m/$734m", "中", "竞争、CRS/感染", "S10"],
  ["双抗", "药品", "美股", "Regeneron", "REGN", "CD20xCD3等双抗平台", "Veloci-Bi平台与临床管线", "低", "临床竞争与组合用药", "S47"],
  ["双抗", "药品", "港股", "康方生物", "9926.HK", "ivonescimab (PD-1/VEGF)", "中国双抗商业化/出海代表", "低", "全球临床验证与伙伴执行", "S48"],
  ["双抗", "药品", "港股", "信达生物", "1801.HK", "PD-1/IL-2等多抗体布局", "国内综合肿瘤与自免平台", "低", "商业化效率、管线取舍", "S49"],
  ["双抗", "药品", "A股", "百利天恒", "688506.SH", "SI-B001等双抗/双抗ADC", "中国差异化多抗体平台", "中", "临床与资金集中度", "S44"],
  ["CGT", "药品", "美股", "Gilead/Kite", "GILD", "Yescarta/Tecartus", "FY25细胞治疗$1.839bn", "中", "竞争与中心产能", "S09"],
  ["CGT", "药品", "美股", "Bristol Myers Squibb", "BMY", "Breyanzi/Abecma", "FY25收入$1.4bn/$427m", "中", "商业制造、竞争", "S14"],
  ["CGT", "药品", "美股/港股", "Johnson & Johnson / Legend", "JNJ / LEGN", "Carvykti", "FY25 $1.887bn，+96%", "中", "产能、供应链、安全性", "S10"],
  ["CGT", "药品", "美股", "Vertex / CRISPR Therapeutics", "VRTX / CRSP", "Casgevy", "首批CRISPR基因编辑商业化", "中", "前处理、价格、收治能力", "S13,S38"],
  ["CGT", "药品", "美股", "Sarepta", "SRPT", "Elevidys/PMO", "FY25商业化基因治疗暴露", "低", "重大安全性/监管风险", "S50"],
  ["CGT", "药品", "A股", "复星医药", "600196.SH", "复星凯特CAR-T", "国内商业化CAR-T平台", "中", "价格/准入、治疗中心", "S51"],
  ["CGT", "CDMO", "A股", "和元生物", "688238.SH", "质粒/病毒载体/CGT CRO-CDMO", "国内纯CGT外包标的", "很高", "负毛利CDMO、亏损与利用率", "S29"],
  ["mRNA", "药品", "美股", "Moderna", "MRNA", "Spikevax/mNEXSPIKE/mRESVIA及肿瘤疫苗", "纯mRNA商业化平台", "中", "疫苗需求收缩、现金消耗", "S15"],
  ["mRNA", "药品", "美股", "BioNTech", "BNTX", "Comirnaty+肿瘤mRNA/免疫组合", "FY25收入€2.87bn，现金€17.2bn", "中", "疫苗下滑、肿瘤管线转化", "S16"],
  ["mRNA", "药品", "美股", "Pfizer", "PFE", "Comirnaty商业化伙伴", "全球商业化与供应链", "中", "疫苗季节性与推荐政策", "S36"],
  ["mRNA", "药品", "A股", "沃森生物", "300142.SZ", "mRNA疫苗合作管线", "国内疫苗平台参与者", "低", "商业化验证与政策", "S52"],
  ["mRNA", "药品", "A股/H股", "康希诺", "688185.SH", "吸入疫苗平台及mRNA布局", "国内疫苗技术平台", "低", "需求周期、持续亏损", "S53"],
  ["多赛道", "临床CRO/数据", "美股", "IQVIA", "IQV", "全球临床+真实世界数据", "R&DS backlog $32.7bn", "很高", "取消率、数据合规", "S21"],
  ["多赛道", "临床CRO", "美股", "Medpace", "MEDP", "中小biotech临床全服务", "FY25收入+20%，backlog $3.03bn", "高", "客户融资与估值", "S22"],
  ["多赛道", "工具/CDMO/CRO", "美股", "Thermo Fisher", "TMO", "仪器、耗材、PPD、Patheon", "FY25收入$44.56bn；一体化平台", "很高", "广泛性稀释纯赛道弹性", "S23"],
  ["多赛道", "工具/耗材", "美股", "Danaher", "DHR", "Cytiva/Pall/Aldevron/Abcam", "FY25 Biotechnology收入$7.293bn，core +6.5%", "很高", "生命科学周期与估值", "S24"],
  ["多赛道", "生物工艺耗材", "美股", "Repligen", "RGEN", "过滤/层析/流体管理", "FY25收入$738m，+16%", "很高", "高估值、客户库存", "S25"],
  ["注射制剂", "包材/给药", "美股", "West Pharma", "WST", "高价值胶塞/给药装置", "FY25收入$3.074bn；肥胖/糖尿病合同制造增长", "高", "客户集中、产品结构", "S26"],
  ["mRNA/CGT", "核酸试剂", "美股", "Maravai/TriLink", "MRVI", "CleanCap、mRNA酶/核酸原料", "核酸工具纯度高", "很高", "疫苗订单集中、收入骤降", "S27"],
  ["多赛道", "层析耗材", "A股", "纳微科技", "688690.SH", "色谱填料/层析介质", "FY25收入9.25亿，毛利率71.3%", "很高", "高估值、国产替代验证", "S31"],
  ["mRNA/CGT", "PEG原料", "A股", "键凯科技", "688356.SH", "药用PEG及衍生物", "FY25收入3.18亿，+40.1%", "高", "客户/产品集中、估值", "S28"],
];
players.getRange(`A4:J${3 + playerRows.length}`).values = playerRows;
body(players, `A4:J${3 + playerRows.length}`);
players.getRange(`A4:A${3 + playerRows.length}`).format.font = { bold: true, color: COLORS.navy };
players.getRange(`G4:G${3 + playerRows.length}`).format.fill = "#F8FAFC";
setWidths(players, { A: 12, B: 17, C: 14, D: 26, E: 18, F: 40, G: 40, H: 18, I: 36, J: 15 });
players.freezePanes.freezeRows(3);
players.freezePanes.freezeColumns(2);

// -------------------- 卖铲子映射 --------------------
title(shovels, "A1:J1", "新技术的“卖铲子”：工艺节点、标的与订单传导");
shovels.getRange("A3:J3").values = [["赛道", "工艺节点", "铲子类型", "A股", "美股", "港股/其他头部", "为什么受益", "订单传导", "需验证的KPI", "主要反例/风险"]];
header(shovels, "A3:J3");
const shovelRows = [
  ["GLP-1", "多肽API工艺与放大", "多肽CDMO/API", "药明康德、凯莱英、ST诺泰", "Thermo Fisher（平台）", "Bachem、PolyPeptide", "大批量多肽合成对纯化、收率与成本高度敏感", "管线后期化→产能预留→商业批", "TIDES/多肽收入、在手订单、反应釜容积", "一次性大订单与过度扩产"],
  ["GLP-1", "分离纯化", "层析填料/设备", "纳微科技", "DHR、RGEN、TMO", "Sartorius", "大规模工艺放大增加耗材和优化需求", "原液放量→层析介质复购", "生物医药收入、大客户验证数", "客户库存与国产替代认证慢"],
  ["GLP-1", "无菌制剂/给药", "胶塞、预充针、自注射器", "东富龙、楚天科技（设备）", "WST、STVN", "Gerresheimer", "长期周期注射需要高价值包材与给药器械", "终端量→组件采购量", "HVP组件增长、肥胖/糖尿病合同制造", "客户定制产线利用率"],
  ["ADC", "linker/payload/HPAPI", "高活小分子CDMO", "药明康德、凯莱英、康龙化成", "TMO", "Lonza、Curia", "工艺、封闭、职业暴露与分析要求高", "候选分子→临床批→商业批", "偶联类收入、高活产能利用率", "单个ADC失败、客户内制"],
  ["ADC", "抗体表达+偶联", "ADC XDC/CDMO", "药明康德（多环节）", "TMO", "药明合联、药明生物、Samsung Biologics、Lonza", "两条工艺链合并并要求一致性", "早期项目数→临床后期占比", "项目数、后期项目、在手订单", "平台竞争与价格下行"],
  ["ADC", "分析/纯化/测试", "层析、质谱、生物测试", "纳微科技、百普赛斯、义翘神州", "DHR、RGEN、TMO", "Sartorius", "DAR、游离payload、聚集体等需要高精分析", "候选者增加→分析方法与耗材增加", "耗材复购、新方法订单", "科研试剂价格竞争"],
  ["双抗", "细胞株/上下游工艺", "大分子CDMO", "康龙化成（大分子）", "TMO", "药明生物、Samsung Biologics、Lonza", "主体流程可复用mAb平台，但分子结构影响表达/聚集", "管线数→开发批次→商业规模", "双抗项目数、商业批次", "平台差异小于ADC，单项目溢价有限"],
  ["双抗", "培养基/一次性系统", "耗材与生物反应器", "奥浦迈、东富龙", "DHR、RGEN、TMO", "Sartorius", "所有大分子项目共享质量体系和耗材", "项目数/批次数→耗材复购", "生物工艺有机增长、利用率", "客户去库存"],
  ["CGT", "质粒/病毒载体", "载体CDMO", "和元生物、康龙化成", "TMO、CRL", "Lonza、Catalent(已私有化)", "载体得率、稳定性、分析与GMP复杂", "IND项目数→工艺开发→GMP批", "CDMO利用率、负毛利是否收窄", "全球产能过剩、项目取消"],
  ["CGT", "细胞处理/冷链", "自动化、冻存、样本管理", "东富龙（设备）", "AZTA、BLFS", "Miltenyi、Sartorius", "自体细胞治疗需要链路身份、冷链与自动处理", "治疗患者数→一次性耗材/冷链次数", "每患者耗材、设备装机数", "自体工艺被通用型/体内编辑替代"],
  ["CGT", "分析/基因测序", "NGS、质控试剂", "百普赛斯、诺唯赞", "ILMN、DHR、TMO", "QIAGEN", "整合位点、脱靶、复制型病毒等需高强度检测", "管线数→方法开发/放行测试", "分析收入、试剂复购", "方法标准化后价格下行"],
  ["mRNA", "帽类似物/核苷酸/酶", "核酸试剂", "诺唯赞", "MRVI、DHR、TMO", "Merck KGaA", "mRNA转录、修饰、封帽直接决定表达与免疫原性", "临床批次→GMP试剂需求", "TriLink基础业务、高批量订单占比", "大流行订单消失、客户集中"],
  ["mRNA", "LNP/药用PEG", "脂质与PEG原料", "键凯科技", "无纯美股龙头", "Croda、Evonik、Merck KGaA", "递送是mRNA的成药性核心", "候选者→工艺锁定→商业批", "商业化产品订单、海外客户收入", "专利设计绕开、客户内制"],
  ["mRNA", "纯化/无菌灌装", "过滤、CDMO、包材", "纳微科技、东富龙", "RGEN、TMO、WST", "Sartorius、Lonza", "从核酸原液到LNP制剂需要多步放大与无菌保障", "商业批次→过滤/包材复购", "有机增长、灌装线利用率", "疫苗产能过剩"],
];
shovels.getRange(`A4:J${3 + shovelRows.length}`).values = shovelRows;
body(shovels, `A4:J${3 + shovelRows.length}`);
shovels.getRange(`A4:A${3 + shovelRows.length}`).format.font = { bold: true, color: COLORS.navy };
shovels.getRange(`G4:I${3 + shovelRows.length}`).format.fill = "#F8FAFC";
setWidths(shovels, { A: 11, B: 27, C: 27, D: 33, E: 26, F: 34, G: 49, H: 34, I: 35, J: 39 });
shovels.freezePanes.freezeRows(3);
shovels.freezePanes.freezeColumns(2);

// -------------------- 公司量化 --------------------
title(companies, "A1:W1", "卖铲子/外包公司量化筛选（FY2025 + 2026-08-07 价格）");
companies.getRange("A2:W2").merge();
companies.getRange("A2:W2").values = [["P/E = 当前价格 / EPS口径（A股为FY2025；美股为实时数据源TTM，WST为FY2025）。分数是可比的筛选工具，不是目标价。"]];
companies.getRange("A2:W2").format = { fill: COLORS.pale, font: { color: COLORS.slate, italic: true }, wrapText: true };
companies.getRange("A4:W4").values = [[
  "市场", "公司", "代码", "主要定位", "主要受益赛道", "币种", "FY2025收入(十亿)", "收入增长", "利润/利润率口径", "EPS口径", "当前价格", "P/E", "市值(USD bn)", "增长分", "盈利质量", "壁垒", "分散度", "估值分", "风险扣分", "综合分", "观察要点", "评分含义", "来源ID"
]];
header(companies, "A4:W4");
const companyRows = [
  ["A股", "药明康德", "603259.SH", "CRDMO/TIDES", "GLP-1、ADC、多赛道", "CNY", 43.42, 0.214, "经调整净利率34.4%", 5.16, 154.82, null, null, null, 5, 5, 5, null, 5, null, "TIDES +96%；持续经营订単+28.8%", "高质量平台，需对地缘风险打折", "S17,S35"],
  ["A股", "凯莱英", "002821.SZ", "CDMO/化学大分子", "GLP-1、ADC", "CNY", 6.67, 0.1491, "归母净利率17.0%", 3.16, 172.26, null, null, null, 4, 4, 4, null, 3, null, "2026收入指引+19%-22%", "转向新分子的CDMO龙头，关注估值与爬坡", "S20,S35"],
  ["A股", "康龙化成", "300759.SZ", "全流程CRO/CDMO", "ADC、CGT、多赛道", "CNY", 14.095, 0.1482, "经调整净利率12.9%", 0.9443, 45.28, null, null, null, 4, 4, 5, null, 3, null, "新签订单+14%；跨板块客户收入占77%", "平台多样性优，估值与海外执行是变量", "S19,S35"],
  ["A股", "泰格医药", "300347.SZ", "临床CRO", "多赛道", "CNY", 6.833, 0.0348, "扣非净利率5.2%", 1.04, 54.35, null, null, null, 3, 4, 3, null, 2, null, "收入恢复弱，扣非利润-58.5%", "临床资产轻，但当前盈利质量需验证", "S18,S35"],
  ["A股", "ST诺泰", "688076.SH", "多肽API/CDMO", "GLP-1", "CNY", 1.943, 0.1957, "扣非净利率18.8%", 1.06, 28.29, null, null, null, 4, 4, 2, null, 25, null, "收入+19.6%但扣非利润-10.2%；ST", "技术暴露纯，但ST/治理风险优先级最高", "S32,S35"],
  ["A股", "纳微科技", "688690.SH", "层析填料/介质", "GLP-1、ADC、双抗、CGT", "CNY", 0.925, 0.1818, "扣非净利率12.7%；毛利71.3%", 0.3394, 42.00, null, null, null, 4, 4, 5, null, 2, null, "色谱/层析收入+22.6%", "高毛利国产替代铲子，但估值较高", "S31,S35"],
  ["A股", "键凯科技", "688356.SH", "药用PEG", "mRNA、CGT、长效制剂", "CNY", 0.318, 0.4013, "扣非净利率15.0%", 1.02, 74.50, null, null, null, 4, 4, 3, null, 3, null, "海外商业化客户带动收入+40%", "高纯度原料暴露，需防客户集中与估值", "S28,S35"],
  ["A股", "和元生物", "688238.SH", "CGT CRO/CDMO", "CGT", "CNY", 0.268, 0.0788, "归母净利率-87.8%；CDMO毛亏", -0.36, 6.93, null, null, null, 1, 3, 3, null, 10, null, "CDMO收入1.356亿、成本2.430亿", "有稀缺性但尚未证明产能经济性", "S29,S35"],
  ["美股", "IQVIA", "IQV", "临床CRO+数据", "多赛道", "USD", 16.31, 0.059, "GAAP净利率8.3%", 8.04, 235.59, null, 39.44, null, 4, 5, 5, null, 1, null, "R&DS backlog $32.7bn；book-to-bill 1.12x", "最宽护城河临床平台之一，增速温和", "S21,S36"],
  ["美股", "Medpace", "MEDP", "临床CRO", "多赛道", "USD", 2.530, 0.20, "GAAP净利率17.8%", 17.04, 587.95, null, 16.79, null, 5, 4, 4, null, 2, null, "FY25收入+20%；净利率仍较高", "成长最快的临床CRO样本之一，估值不低", "S22,S36"],
  ["美股", "Thermo Fisher", "TMO", "工具+CRO+CDMO", "全部", "USD", 44.556, 0.039, "多分部盈利；LSS营业利润率36.3%", 18.59, 586.35, null, 217.54, null, 5, 5, 5, null, 1, null, "一体化Accelerator获得订单", "最分散的卖铲子，弹性低于纯标的但韧性强", "S23,S36"],
  ["美股", "Danaher", "DHR", "生物工艺/生命科学", "双抗、ADC、CGT、mRNA", "USD", 24.568, 0.029, "GAAP营业利润率19.1%", 5.64, 202.52, null, 143.30, null, 5, 5, 5, null, 1, null, "Biotechnology core sales +6.5%", "Cytiva/Pall/Aldevron组合跨模态复用", "S24,S36"],
  ["美股", "Repligen", "RGEN", "生物工艺耗材", "双抗、ADC、CGT、mRNA", "USD", 0.738, 0.16, "GAAP盈利，但利润基数较低", 0.74, 163.43, null, 9.26, null, 3, 5, 5, null, 4, null, "FY25收入+16%，有机非COVID +16%", "高弹性耗材，但GAAP P/E极高", "S25,S36"],
  ["美股", "West Pharma", "WST", "注射包材/给药", "GLP-1、ADC、双抗", "USD", 3.074, 0.063, "自由现金流$469m", 6.79, 353.41, null, 25.74, null, 5, 5, 3, null, 2, null, "HVP组件Q4有机+15.1%", "高质量给药铲子，受益GLP-1但估值较高", "S26,S36"],
  ["美股", "Maravai", "MRVI", "mRNA/核酸试剂", "mRNA、CGT", "USD", 0.186, -0.284, "Adjusted EBITDA -$31m", -0.71, 5.88, null, 0.86, null, 1, 4, 3, null, 10, null, "TriLink收入-39%；2026指引恢复增长", "技术铲子纯度高，但尚在大流行去产能周期", "S27,S36"],
  ["美股", "Charles River", "CRL", "临床前+制造", "多赛道、CGT", "USD", 4.02, -0.009, "GAAP亏损；Non-GAAP营业利润率19.8%", -4.80, 264.57, null, 12.70, null, 2, 4, 4, null, 5, null, "DSA收入-2%；制造分部减值", "底层需求有壁垒，但制造资产质量需修复", "S34,S36"],
];
companies.getRange(`A5:W${4 + companyRows.length}`).values = companyRows;
body(companies, `A5:W${4 + companyRows.length}`);
for (let r = 5; r <= 4 + companyRows.length; r++) {
  companies.getRange(`L${r}`).formulas = [[`=IF(J${r}>0,K${r}/J${r},"")`]];
  companies.getRange(`N${r}`).formulas = [[`=IF(H${r}>='评分模型'!$A$34,'评分模型'!$B$34,IF(H${r}>='评分模型'!$A$35,'评分模型'!$B$35,IF(H${r}>='评分模型'!$A$36,'评分模型'!$B$36,IF(H${r}>='评分模型'!$A$37,'评分模型'!$B$37,'评分模型'!$B$38))))`]];
  companies.getRange(`R${r}`).formulas = [[`=IF(OR(L${r}="",L${r}<=0),1,IF(L${r}<='评分模型'!$A$42,'评分模型'!$B$42,IF(L${r}<='评分模型'!$A$43,'评分模型'!$B$43,IF(L${r}<='评分模型'!$A$44,'评分模型'!$B$44,IF(L${r}<='评分模型'!$A$45,'评分模型'!$B$45,1)))))`]];
  companies.getRange(`T${r}`).formulas = [[`=MAX(0,20*SUMPRODUCT(N${r}:R${r},'评分模型'!$B$26:$B$30)-S${r})`]];
}
companies.getRange(`G5:G${4 + companyRows.length}`).format.numberFormat = "0.000";
companies.getRange(`H5:H${4 + companyRows.length}`).format.numberFormat = "0.0%";
companies.getRange(`J5:K${4 + companyRows.length}`).format.numberFormat = "0.00";
companies.getRange(`L5:L${4 + companyRows.length}`).format.numberFormat = "0.0x";
companies.getRange(`M5:M${4 + companyRows.length}`).format.numberFormat = "0.00";
companies.getRange(`N5:R${4 + companyRows.length}`).format.horizontalAlignment = "center";
companies.getRange(`S5:S${4 + companyRows.length}`).format = { font: { color: COLORS.inputBlue }, fill: "#EFF6FF", horizontalAlignment: "center" };
companies.getRange(`O5:Q${4 + companyRows.length}`).format = { font: { color: COLORS.inputBlue }, fill: "#EFF6FF", horizontalAlignment: "center" };
companies.getRange(`T5:T${4 + companyRows.length}`).format.numberFormat = "0.0";
companies.getRange(`T5:T${4 + companyRows.length}`).conditionalFormats.add("colorScale", { colors: ["#FEE2E2", "#FEF3C7", "#DCFCE7"] });
companies.getRange(`S5:S${4 + companyRows.length}`).conditionalFormats.add("colorScale", { colors: ["#DCFCE7", "#FEF3C7", "#FEE2E2"] });
setWidths(companies, { A: 10, B: 19, C: 15, D: 25, E: 29, F: 9, G: 15, H: 13, I: 30, J: 11, K: 13, L: 10, M: 15, N: 10, O: 11, P: 10, Q: 10, R: 10, S: 12, T: 11, U: 44, V: 46, W: 17 });
companies.freezePanes.freezeRows(4);
companies.freezePanes.freezeColumns(3);

// -------------------- 总览 --------------------
title(overview, "A1:L1", "医药行业基础架构与新技术赛道量化总览");
overview.getRange("A2:L2").merge();
overview.getRange("A2:L2").values = [["截至 2026-08-07｜核心范围：CRO/CMO/CDMO/CRDMO，GLP-1、ADC、双抗、CGT、mRNA，A股+美股（港股作必要补充）"]];
overview.getRange("A2:L2").format = { fill: COLORS.pale, font: { color: COLORS.slate, italic: true }, horizontalAlignment: "center" };

section(overview, "A4:L4", "一页结论");
overview.getRange("A5:C5").merge();
overview.getRange("D5:F5").merge();
overview.getRange("G5:I5").merge();
overview.getRange("J5:L5").merge();
overview.getRange("A5:C5").values = [["近端赛道首选"]];
overview.getRange("D5:F5").values = [["卖铲子首选"]];
overview.getRange("G5:I5").values = [["最大误区"]];
overview.getRange("J5:L5").values = [["组合思路"]];
for (const r of ["A5:C5", "D5:F5", "G5:I5", "J5:L5"]) overview.getRange(r).format = { fill: COLORS.lightSlate, font: { bold: true, color: COLORS.navy }, horizontalAlignment: "center" };
overview.getRange("A6:C7").merge(); overview.getRange("A6:C7").values = [["GLP-1：终端收入与产能需求最强"]];
overview.getRange("D6:F7").merge(); overview.getRange("D6:F7").values = [["ADC：偶联/linker/payload/分析最容易获得行业β"]];
overview.getRange("G6:I7").merge(); overview.getRange("G6:I7").values = [["CGT项目数最多≠上游一定赚钱；利用率是关键"]];
overview.getRange("J6:L7").merge(); overview.getRange("J6:L7").values = [["核心平台+细分高弹性+技术期权，避免单管线押注"]];
for (const r of ["A6:C7", "D6:F7", "G6:I7", "J6:L7"]) overview.getRange(r).format = { fill: COLORS.white, font: { bold: true, color: COLORS.black, size: 12 }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "outside", style: "thin", color: "#CBD5E1" } };

section(overview, "A9:L9", "赛道评分（0-100，权重可在“评分模型”页调整）");
overview.getRange("A10:D10").values = [["赛道", "创新药分", "卖铲子分", "综合分"]];
header(overview, "A10:D10");
for (let i = 0; i < 5; i++) {
  const r = 11 + i;
  const src = 5 + i;
  overview.getRange(`A${r}:D${r}`).formulas = [[`='赛道评分'!A${src}`, `='赛道评分'!I${src}`, `='赛道评分'!O${src}`, `='赛道评分'!P${src}`]];
}
body(overview, "A11:D15");
overview.getRange("A11:D15").format.font = { color: COLORS.linkGreen };
overview.getRange("B11:D15").format.numberFormat = "0.0";
overview.getRange("B11:D15").conditionalFormats.add("colorScale", { colors: ["#FEE2E2", "#FEF3C7", "#DCFCE7"] });
const scoreChart = overview.charts.add("bar", overview.getRange("A10:D15"));
scoreChart.title = "赛道综合吸引力（0-100）";
scoreChart.hasLegend = true;
scoreChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 10 } };
scoreChart.yAxis = { numberFormatCode: "0", min: 0, max: 100 };
scoreChart.setPosition("F10", "L24");

section(overview, "A17:E17", "业务模式投资观察");
overview.getRange("A18:E18").values = [["模式", "可见度", "资本强度", "锁定性", "要看什么"]];
header(overview, "A18:E18");
overview.getRange("A19:E23").values = [
  ["CRO", "中", "低-中", "中", "Backlog、book-to-bill、取消率"],
  ["CMO", "中", "高", "中", "产能利用率、良率、客户集中"],
  ["CDMO", "高", "高", "高", "后期/商业化项目占比"],
  ["CRDMO", "很高", "中-高", "很高", "转化率、在手订单、交叉销售"],
  ["工具/耗材", "高", "中", "高", "有机增长、复购、去库存进程"],
];
body(overview, "A19:E23");

section(overview, "A25:L25", "量化信号：不能忽略的六个数");
overview.getRange("A26:L29").values = [
  ["GLP-1商业收入", "$36.5bn", "Lilly两大tirzepatide品牌FY25", "ADC销售验证", "$3.575bn", "Enhertu 9M25合并销售", "双抗销售验证", "CHF8.856bn", "Hemlibra+Vabysmo FY25", "CGT分化", "+96% / -7%", "Carvykti vs Gilead细胞治疗"],
  ["CRO订单", "$32.7bn", "IQVIA R&DS backlog", "CRDMO订单", "¥58.0bn", "药明康德持续经营在手订单", "多肽高增", "+96%", "药明TIDES收入增速", "mRNA去产能", "-39%", "Maravai TriLink FY25"],
  ["A股层析铲子", "71.3%", "纳微科技FY25毛利率", "CGT产能警报", "负毛利", "和元CGT CDMO成本>收入", "美股工具恢复", "+16%", "Repligen FY25收入", "注射包材", "+15.1%", "West HVP Q4有机增长"],
  ["赛道热度", "2,061", "CGT活跃干预性试验", "GLP-1热度", "777", "活跃干预性试验", "ADC热度", "565", "活跃干预性试验", "mRNA热度", "95", "活跃干预性试验"],
];
body(overview, "A26:L29");
overview.getRange("A26:L29").format.borders = { preset: "all", style: "thin", color: "#E2E8F0" };
overview.getRange("A26:L29").format.rowHeight = 40;

overview.getRange("A31:L33").merge();
overview.getRange("A31:L33").values = [["使用方法：先在“评分模型”页调整权重，再看“赛道评分”的组合排名；通过“卖铲子映射”找到受益环节，最后用“公司量化”检查增长、盈利、估值与风险是否一致。本工作簿不构成个股投资建议。"]];
overview.getRange("A31:L33").format = { fill: COLORS.lightAmber, font: { color: "#78350F", italic: true }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: "#F59E0B" } };
setWidths(overview, { A: 17, B: 13, C: 24, D: 17, E: 16, F: 24, G: 17, H: 14, I: 23, J: 17, K: 14, L: 27 });
overview.freezePanes.freezeRows(4);

// -------------------- 来源与口径 --------------------
title(sources, "A1:G1", "来源、口径与局限");
sources.getRange("A3:G3").values = [["来源ID", "主题", "数值/口径", "期间/时点", "来源名称", "URL", "备注"]];
header(sources, "A3:G3");
const sourceRows = [
  ["S01", "ClinicalTrials.gov API", "API v2；数据工作日更新", "2026-08-07 09:00 ET dataTimestamp", "ClinicalTrials.gov API documentation", "https://clinicaltrials.gov/data-api", "试验数是检索结果，不代表成功概率"],
  ["S02", "GLP-1活跃试验", 777, "2026-08-07", "ClinicalTrials.gov API", "https://clinicaltrials.gov/api/v2/studies?query.term=AREA%5BStudyType%5DINTERVENTIONAL%20AND%20%22GLP-1%22&filter.overallStatus=RECRUITING%7CNOT_YET_RECRUITING%7CACTIVE_NOT_RECRUITING%7CENROLLING_BY_INVITATION&pageSize=1&countTotal=true", "搜索词口径会影响结果"],
  ["S03", "ADC活跃试验", 565, "2026-08-07", "ClinicalTrials.gov API", "https://clinicaltrials.gov/api/v2/studies?query.term=AREA%5BStudyType%5DINTERVENTIONAL%20AND%20%22antibody-drug%20conjugate%22&filter.overallStatus=RECRUITING%7CNOT_YET_RECRUITING%7CACTIVE_NOT_RECRUITING%7CENROLLING_BY_INVITATION&pageSize=1&countTotal=true", "同上"],
  ["S04", "双抗活跃试验", 485, "2026-08-07", "ClinicalTrials.gov API", "https://clinicaltrials.gov/api/v2/studies?query.term=AREA%5BStudyType%5DINTERVENTIONAL%20AND%20%22bispecific%20antibody%22&filter.overallStatus=RECRUITING%7CNOT_YET_RECRUITING%7CACTIVE_NOT_RECRUITING%7CENROLLING_BY_INVITATION&pageSize=1&countTotal=true", "同上"],
  ["S05", "CGT活跃试验", 2061, "2026-08-07", "ClinicalTrials.gov API", "https://clinicaltrials.gov/api/v2/studies?query.term=AREA%5BStudyType%5DINTERVENTIONAL%20AND%20(%22cell%20therapy%22%20OR%20%22gene%20therapy%22)&filter.overallStatus=RECRUITING%7CNOT_YET_RECRUITING%7CACTIVE_NOT_RECRUITING%7CENROLLING_BY_INVITATION&pageSize=1&countTotal=true", "口径较广，分数用对数归一化"],
  ["S06", "mRNA活跃试验", 95, "2026-08-07", "ClinicalTrials.gov API", "https://clinicaltrials.gov/api/v2/studies?query.term=AREA%5BStudyType%5DINTERVENTIONAL%20AND%20(%22mRNA%20vaccine%22%20OR%20%22mRNA%20therapeutic%22)&filter.overallStatus=RECRUITING%7CNOT_YET_RECRUITING%7CACTIVE_NOT_RECRUITING%7CENROLLING_BY_INVITATION&pageSize=1&countTotal=true", "同上"],
  ["S07", "Lilly GLP-1销售", "Mounjaro $22.965bn; Zepbound $13.542bn", "FY2025", "Eli Lilly 2025 10-K", "https://investor.lilly.com/static-files/70c644f5-d894-471a-a4c5-6464848ca393", "公司10-K"],
  ["S08", "Novo GLP-1销售", "Ozempic DKK127.089bn; Wegovy DKK79.106bn; Rybelsus DKK23.301bn", "FY2025", "Novo Nordisk Annual Report 2025", "https://annualreport.novonordisk.com/2025/", "公司年报"],
  ["S09", "Gilead ADC/CGT", "Trodelvy $1.397bn; Cell Therapy $1.839bn", "FY2025", "Gilead FY2025 results", "https://investors.gilead.com/news/news-details/2026/Gilead-Sciences-Announces-Fourth-Quarter-and-Full-Year-2025-Financial-Results/default.aspx", "公司业绩"],
  ["S10", "JNJ CGT/双抗", "Carvykti $1.887bn; Tecvayli $670m; Talvey $463m; Rybrevant/Lazcluze $734m", "FY2025", "Johnson & Johnson 2025 Annual Report", "https://www.jnj.com/download/johnson-johnson-2025-annual-report", "公司年报"],
  ["S11", "Roche双抗", "Hemlibra CHF4.754bn; Vabysmo CHF4.102bn", "FY2025", "Roche FY2025 results", "https://www.roche.com/investors/updates/inv-update-2026-01-29", "公司业绩"],
  ["S12", "Amgen双抗", "Blincyto $1.559bn", "FY2025", "Amgen FY2025 results", "https://investors.amgen.com/news-releases/news-release-details/amgen-reports-fourth-quarter-and-full-year-2025-financial", "公司业绩"],
  ["S13", "FDA CGT获批产品", "FDA OTP清单", "当前", "FDA Approved Cellular and Gene Therapy Products", "https://www.fda.gov/vaccines-blood-biologics/cellular-gene-therapy-products/approved-cellular-and-gene-therapy-products", "官方获批清单"],
  ["S14", "BMS CGT", "Breyanzi $1.4bn; Abecma $427m", "FY2025", "Bristol Myers Squibb key facts", "https://www.bms.com/investors/financial-reporting/key-facts.html", "公司业绩"],
  ["S15", "Moderna", "总收入$1.9bn，-40%", "FY2025", "Moderna FY2025 results / SEC", "https://www.sec.gov/Archives/edgar/data/1682852/000168285226000015/exhibit9912025q4pressrelea.htm", "大部分为COVID疫苗"],
  ["S16", "BioNTech", "收入€2.870bn；现金/证券€17.2bn", "FY2025", "BioNTech FY2025 results", "https://investors.biontech.de/news-releases/news-release-details/biontech-announces-fourth-quarter-and-full-year-2025-financial/", "收入含BMS合作影响"],
  ["S17", "药明康德", "持续经营收入434.2亿(+21.4%); TIDES 113.7亿(+96%); 订单580亿", "FY2025", "WuXi AppTec FY2025 results", "https://www.prnasia.com/story/526325-1.shtml", "公司新闻稿/已审计年报数据"],
  ["S18", "泰格医药", "收入68.33亿(+3.48%); 扣非净利-58.47%", "FY2025", "Tigermed 2025 Annual Report", "https://tigermedgrp.com/files/23a461008465697e3b38eaff54a1d9bb.pdf", "已审计年报"],
  ["S19", "康龙化成", "收入140.95亿(+14.82%); 扣非净利+38.85%", "FY2025", "Pharmaron 2025 Annual Report", "https://vip.stock.finance.sina.com.cn/corp/view/vCB_AllBulletinDetail.php?id=12042867", "二次镜像，源文件为公司年报"],
  ["S20", "凯莱英", "收入66.70亿(+14.91%); 归母净利+19.35%", "FY2025", "Asymchem 2025 Annual Report", "https://money.finance.sina.com.cn/corp/view/vCB_AllBulletinDetail.php?id=12042560", "二次镜像，源文件为公司年报"],
  ["S21", "IQVIA", "收入$16.31bn; R&DS backlog $32.7bn; B2B 1.12x", "FY2025", "IQVIA FY2025 results", "https://ir.iqvia.com/press-releases/press-release-details/2026/IQVIA-Reports-Fourth-Quarter-and-Full-Year-2025-Results-Issues-Full-Year-2026-Guidance/default.aspx", "公司业绩"],
  ["S22", "Medpace", "收入$2.530bn(+20%); backlog $3.027bn", "FY2025", "Medpace FY2025 results", "https://investor.medpace.com/news-releases/news-release-details/medpace-holdings-inc-reports-fourth-quarter-and-full-year-2025", "公司业绩"],
  ["S23", "Thermo Fisher", "收入$44.556bn(+4%); LSS $10.374bn", "FY2025", "Thermo Fisher FY2025 results", "https://ir.thermofisher.com/investors/news-events/news/news-details/2026/Thermo-Fisher-Scientific-Reports-Fourth-Quarter-and-Full-Year-2025-Results/default.aspx", "公司业绩"],
  ["S24", "Danaher", "收入$24.568bn; Biotechnology $7.293bn", "FY2025", "Danaher FY2025 results", "https://investors.danaher.com/2026-01-28-Danaher-Reports-Fourth-Quarter-and-Full-Year-2025-Results", "公司业绩"],
  ["S25", "Repligen", "收入$738m(+16%)", "FY2025", "Repligen FY2025 results", "https://investors.repligen.com/financials/quarterly-results/default.aspx", "公司业绩"],
  ["S26", "West Pharma", "收入$3.074bn(+6.3%); FCF $468.9m", "FY2025", "West FY2025 results", "https://investor.westpharma.com/news-releases/news-release-details/west-reports-fourth-quarter-and-full-year-2025-results", "公司业绩"],
  ["S27", "Maravai", "收入$185.7m; TriLink $119.8m(-39%)", "FY2025", "Maravai FY2025 results", "https://investors.maravai.com/news-events/press-releases/detail/123/maravai-lifesciences-reports-fourth-quarter-and-full-year-2025-financial-results", "公司业绩"],
  ["S28", "键凯科技", "收入3.18亿(+40.13%)", "FY2025", "Jenkem 2025 Annual Report", "https://money.finance.sina.com.cn/corp/view/vCB_AllBulletinDetail.php?id=12247133", "二次镜像"],
  ["S29", "和元生物", "收入2.677亿(+7.88%); 净亏2.351亿", "FY2025", "Obio 2025 Annual Report", "https://money.finance.sina.com.cn/corp/view/vCB_AllBulletinDetail.php?id=12117811", "二次镜像"],
  ["S30", "ST诺泰", "收入19.43亿(+19.57%); 扣非净利-10.21%", "FY2025", "Sinopep-Allsino 2025 Annual Report", "https://money.finance.sina.com.cn/corp/view/vCB_AllBulletinDetail.php?id=12283906", "当前证券简称已带ST"],
  ["S31", "纳微科技", "收入9.25亿(+18.18%); 毛利率71.31%", "FY2025", "Nano-Micro 2025 Annual Report", "https://money.finance.sina.com.cn/corp/view/vCB_AllBulletinDetail.php?id=12094244", "二次镜像"],
  ["S32", "诺泰基本每股收益", "基本EPS 1.06元", "FY2025", "Sinopep-Allsino company filing", "https://dataclouds.cninfo.com.cn/shgonggao/hsomarket/2026/20260518/88d258457fe64fb7825f8c8c6070e77d.PDF", "用于粗略P/E"],
  ["S33", "A股收盘价", "603259/300759/300347/002821/688076/688238/688356/688690", "2026-08-07", "Sina real-time quote endpoint", "https://hq.sinajs.cn/list=sh603259,sz300759,sz300347,sz002821,sh688076,sh688238,sh688356,sh688690", "需Referer；用于价格/EPS粗略P/E"],
  ["S34", "Charles River", "收入$4.02bn(-0.9%); GAAP loss; adjusted op margin 19.8%", "FY2025", "Charles River FY2025 results", "https://ir.criver.com/news-releases/news-release-details/charles-river-laboratories-announces-fourth-quarter-and-full-14", "公司业绩"],
  ["S35", "A股价格页", "价格核验", "2026-08-07", "Sina Finance", "https://finance.sina.com.cn/realstock/company/sh603259/nc.shtml", "其他A股以对应代码页面核验"],
  ["S36", "美股实时价格/估值", "价格、市值、P/E", "2026-08-07", "Market quote pages", "https://finance.yahoo.com/quote/TMO/", "其他美股以对应ticker页面核验"],
  ["S37", "AbbVie", "ADC平台与FY2025业绩", "FY2025", "AbbVie FY2025 results", "https://news.abbvie.com/2026-02-04-AbbVie-Reports-Full-Year-and-Fourth-Quarter-2025-Financial-Results", "公司业绩"],
  ["S38", "Vertex", "Casgevy及FY2025业绩", "FY2025", "Vertex FY2025 results", "https://investors.vrtx.com/news-releases/news-release-details/vertex-reports-fourth-quarter-and-full-year-2025-financial", "公司业绩"],
  ["S39", "Viking Therapeutics", "VK2735管线", "当前", "Viking investor relations", "https://ir.vikingtherapeutics.com/", "公司IR"],
  ["S40", "恒瑞医药", "公司管线/披露", "当前", "Hengrui official website", "https://www.hengrui.com/", "公司官网"],
  ["S41", "华东医药", "公司管线/披露", "当前", "Huadong Medicine official website", "https://www.eastchinapharm.com/", "公司官网"],
  ["S42", "博瑞医药", "公司管线/披露", "当前", "BrightGene official website", "https://www.bright-gene.com/", "公司官网"],
  ["S43", "科伦药业", "公司管线/披露", "当前", "Kelun official website", "https://www.kelun.com/", "公司官网"],
  ["S44", "百利天恒", "公司管线/披露", "当前", "Baili Pharmaceutical official website", "https://www.baili-pharm.com/", "公司官网"],
  ["S45", "荣昌生物", "公司管线/披露", "当前", "RemeGen official website", "https://www.remegen.com/", "公司官网"],
  ["S46", "药明合联", "偶联药CRDMO", "当前", "WuXi XDC official website", "https://www.wuxixdc.com/", "公司官网"],
  ["S47", "Regeneron", "双抗平台/管线", "当前", "Regeneron pipeline", "https://www.regeneron.com/science/pipeline", "公司官网"],
  ["S48", "康方生物", "ivonescimab等管线", "当前", "Akeso official website", "https://www.akesobio.com/", "公司官网"],
  ["S49", "信达生物", "多抗体管线", "当前", "Innovent official website", "https://www.innoventbio.com/", "公司官网"],
  ["S50", "Sarepta", "FY2025业绩/Elevidys", "FY2025", "Sarepta FY2025 results", "https://investorrelations.sarepta.com/news-releases/news-release-details/sarepta-therapeutics-announces-fourth-quarter-and-full-year-2025", "公司业绩"],
  ["S51", "复星医药", "复星凯特CAR-T", "当前", "Fosun Pharma official website", "https://www.fosunpharma.com/", "公司官网"],
  ["S52", "沃森生物", "mRNA合作管线", "当前", "Walvax official website", "https://www.walvax.com/", "公司官网"],
  ["S53", "康希诺", "疫苗平台/管线", "当前", "CanSino official website", "https://www.cansinotech.com/", "公司官网"],
];
sources.getRange(`A4:G${3 + sourceRows.length}`).values = sourceRows;
body(sources, `A4:G${3 + sourceRows.length}`);
sources.getRange(`F4:F${3 + sourceRows.length}`).format.font = { color: COLORS.blue, underline: true };
sources.getRange("C4:C9").format.numberFormat = "#,##0";
setWidths(sources, { A: 11, B: 29, C: 40, D: 24, E: 38, F: 78, G: 46 });
sources.freezePanes.freezeRows(3);
sources.freezePanes.freezeColumns(1);

// -------------------- 检查 --------------------
title(checks, "A1:G1", "模型检查");
checks.getRange("A3:G3").values = [["检查项", "实际值", "期望值", "差异", "容差", "状态", "备注"]];
header(checks, "A3:G3");
checks.getRange("A4:G9").values = [
  ["创新药权重合计", null, 1, null, 0.0001, null, "评分模型 B5:B9"],
  ["卖铲子权重合计", null, 1, null, 0.0001, null, "评分模型 B13:B17"],
  ["组合权重合计", null, 1, null, 0.0001, null, "评分模型 B21:B22"],
  ["公司权重合计", null, 1, null, 0.0001, null, "评分模型 B26:B30"],
  ["赛道分有效范围", null, 1, null, 0, null, "所有综合分均应在0-100"],
  ["公司分有效范围", null, 1, null, 0, null, "所有综合分均应在0-100"],
];
checks.getRange("B4").formulas = [["=SUM('评分模型'!B5:B9)"]];
checks.getRange("B5").formulas = [["=SUM('评分模型'!B13:B17)"]];
checks.getRange("B6").formulas = [["=SUM('评分模型'!B21:B22)"]];
checks.getRange("B7").formulas = [["=SUM('评分模型'!B26:B30)"]];
checks.getRange("B8").formulas = [["=IF(AND(MIN('赛道评分'!I5:I9)>=0,MAX('赛道评分'!P5:P9)<=100),1,0)"]];
checks.getRange("B9").formulas = [[`=IF(AND(MIN('公司量化'!T5:T${4 + companyRows.length})>=0,MAX('公司量化'!T5:T${4 + companyRows.length})<=100),1,0)`]];
for (let r = 4; r <= 9; r++) {
  checks.getRange(`D${r}`).formulas = [[`=B${r}-C${r}`]];
  checks.getRange(`F${r}`).formulas = [[`=IF(ABS(D${r})<=E${r},"OK","CHECK")`]];
}
body(checks, "A4:G9");
checks.getRange("B4:F9").format.horizontalAlignment = "center";
checks.getRange("B4:E9").format.numberFormat = "0.0000";
checks.getRange("F4:F9").conditionalFormats.add("containsText", { text: "OK", format: { fill: COLORS.lightGreen, font: { color: COLORS.green, bold: true } } });
checks.getRange("F4:F9").conditionalFormats.add("containsText", { text: "CHECK", format: { fill: COLORS.lightRed, font: { color: COLORS.red, bold: true } } });
checks.getRange("A11:G13").merge();
checks.getRange("A11:G13").values = [["限制：赛道评分含可编辑定性输入，不是统计预测；不同币种的公司收入不直接比规模；粗略P/E未统一季报时点、经调整口径与净现金，只用于初筛。"]];
checks.getRange("A11:G13").format = { fill: COLORS.lightAmber, font: { color: "#78350F", italic: true }, wrapText: true, verticalAlignment: "center" };
setWidths(checks, { A: 30, B: 14, C: 14, D: 14, E: 12, F: 14, G: 46 });

// Compact body font and row sizing for long tables.
for (const [sheet, range] of [[players, `A4:J${3 + playerRows.length}`], [shovels, `A4:J${3 + shovelRows.length}`], [sources, `A4:G${3 + sourceRows.length}`], [companies, `A5:W${4 + companyRows.length}`]]) {
  sheet.getRange(range).format.rowHeight = 42;
}

// Export.
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

// Compact verification + visual renders of every sheet.
const trackInspect = await workbook.inspect({ kind: "table", range: "赛道评分!A4:P9", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 18 });
const companyInspect = await workbook.inspect({ kind: "table", range: `公司量化!A4:W${4 + companyRows.length}`, include: "values,formulas", tableMaxRows: 8, tableMaxCols: 23 });
const checkInspect = await workbook.inspect({ kind: "table", range: "检查!A3:G9", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 8 });
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" });
console.log("TRACK_INSPECT", trackInspect.ndjson);
console.log("COMPANY_INSPECT", companyInspect.ndjson);
console.log("CHECK_INSPECT", checkInspect.ndjson);
console.log("ERROR_SCAN", errors.ndjson);

const renderSpecs = [
  ["总览", "A1:L33", "overview.png"],
  ["产业架构", "A1:H23", "architecture.png"],
  ["赛道评分", "A1:R9", "tracks.png"],
  ["玩家矩阵", `A1:J${Math.min(18, 3 + playerRows.length)}`, "players_top.png"],
  ["卖铲子映射", `A1:J${3 + shovelRows.length}`, "shovels.png"],
  ["公司量化", `A1:W${4 + companyRows.length}`, "companies.png"],
  ["评分模型", "A1:D46", "model.png"],
  ["来源与口径", "A1:G16", "sources_top.png"],
  ["检查", "A1:G13", "checks.png"],
];
for (const [sheetName, range, fileName] of renderSpecs) {
  const png = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${fileName}`, new Uint8Array(await png.arrayBuffer()));
}

console.log("OUTPUT", outputPath);
