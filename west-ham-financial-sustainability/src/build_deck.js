// Builds the 5-slide deck "West Ham United FC – Financial performance & sustainability".
// Run: NODE_PATH=<dir with pptxgenjs, react, react-dom, react-icons, sharp> node build_deck.js <out.pptx> <pptx-skill-dir>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || path.join(__dirname, "..", "WestHam_Financial_Sustainability.pptx");
const SKILL_DIR = process.argv[3];
const { applyTheme } = require(path.join(SKILL_DIR, "scripts", "apply_theme.js"));

// ---------------------------------------------------------------- theme
const THEME = {
  name: "West Ham Claret",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E1A1D", // near-black text
    lt1: "FFFFFF",
    dk2: "4B1527", // deep claret (dark slides)
    lt2: "F6EEF1", // claret tint (cards on light slides)
    accent1: "7A263A", // West Ham claret
    accent2: "1BB1E7", // West Ham sky blue
    accent3: "A7A9AC", // neutral grey
    accent4: "C0392B", // red  (weak)
    accent5: "E3A21A", // amber (mixed)
    accent6: "2E8B57", // green (strong)
    hlink: "1B7FB0",
    folHlink: "7A263A",
  },
};
const HEX = THEME.colors;
const MUTED = "5F5A5D";
const GRID = "E6E1E3";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "West Ham United FC – Financial performance and sustainability";
pres.subject = "Assignment 9: financial performance and sustainability analysis";
pres.author = "Assignment 9";
const C = pres.SchemeColor;

// ---------------------------------------------------------------- layouts
const FOOTER = "West Ham United FC  ·  Financial performance & sustainability  ·  Assignment 9";
const titlePh = (color) => ({
  placeholder: {
    options: { name: "title", type: "title", x: 0.5, y: 0.32, w: 12.33, h: 0.72, fontSize: 32, bold: true, color, align: "left", valign: "middle", margin: 0 },
    text: "",
  },
});
const subtitlePh = (color) => ({
  placeholder: {
    options: { name: "subtitle", type: "body", x: 0.5, y: 1.04, w: 12.33, h: 0.5, fontSize: 16, color, valign: "top", margin: 0 },
    text: "",
  },
});

pres.defineSlideMaster({
  title: "Content Light",
  background: { color: C.background1 },
  objects: [
    titlePh(C.text2),
    subtitlePh(C.text1),
    { text: { text: FOOTER, options: { x: 0.5, y: 7.08, w: 9, h: 0.28, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: 12.33, y: 7.08, w: 0.5, h: 0.28, fontSize: 10, color: MUTED, align: "right" },
});

pres.defineSlideMaster({
  title: "Content Dark",
  background: { color: C.text2 },
  objects: [
    titlePh(C.background1),
    subtitlePh(C.background2),
    { text: { text: FOOTER, options: { x: 0.5, y: 7.08, w: 9, h: 0.28, fontSize: 10, color: C.background2, margin: 0 } } },
  ],
  slideNumber: { x: 12.33, y: 7.08, w: 0.5, h: 0.28, fontSize: 10, color: C.background2, align: "right" },
});

// ---------------------------------------------------------------- helpers
async function icon(Icon, colorHex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color: "#" + colorHex, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const chartBase = () => ({
  catAxisLabelColor: MUTED,
  valAxisLabelColor: MUTED,
  catAxisLabelFontFace: "+mn-lt",
  valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 11,
  valAxisLabelFontSize: 10,
  dataLabelFontFace: "+mn-lt",
  dataLabelFontSize: 11,
  dataLabelColor: HEX.dk1,
  titleFontFace: "+mn-lt",
  titleFontSize: 13,
  titleColor: HEX.dk1,
  catAxisLineShow: true,
  catAxisLineColor: GRID,
  valGridLine: { style: "none" },
  catGridLine: { style: "none" },
  valAxisHidden: true,
  showTitle: true,
});

function sourceLine(slide, text, dark = false) {
  slide.addText(text, {
    x: 0.5, y: 6.7, w: 12.33, h: 0.32, fontSize: 10, italic: true, margin: 0, valign: "top",
    color: dark ? C.background2 : MUTED, isTextBox: true, objectName: "Source note",
  });
}

function card(slide, x, y, w, h, fill, name) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, objectName: name });
}

// ================================================================ SLIDE 1 – executive summary
async function slide1() {
  pres.addSection({ title: "Executive summary" });
  const s = pres.addSlide({ masterName: "Content Dark", sectionTitle: "Executive summary" });
  s.addText("ASSIGNMENT 9  ·  FINANCIAL PERFORMANCE & SUSTAINABILITY  ·  NOVEMBER 2026", {
    x: 0.5, y: 0.3, w: 12.33, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: C.accent2, margin: 0, isTextBox: true, objectName: "Kicker",
  });
  s.addText("West Ham United: a Premier League cost base without Premier League income", {
    x: 0.5, y: 0.66, w: 12.33, h: 1.25, fontSize: 36, bold: true, fontFace: "+mj-lt", color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: "Deck title",
  });
  s.addText("Is the club financially healthy and sustainable – and what should management prioritise over the next three years?", {
    x: 0.5, y: 1.98, w: 12.33, h: 0.4, fontSize: 16, italic: true, color: C.background2, margin: 0, isTextBox: true, objectName: "Research question",
  });

  // left: answer in brief
  s.addText("The answer in brief", { x: 0.5, y: 2.65, w: 6.9, h: 0.38, fontSize: 18, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "Answer header" });
  const bullets = [
    ["Revenue has stalled: ", "a record £270m in 2023/24 depended on Europe and fell 16% to £228m in 2024/25 – before relegation."],
    ["Spending outran income: ", "wages rose 29% in two years to £176m (77% of revenue vs. 65% Premier League average), on top of c. £570m of player purchases since 2022."],
    ["Profit only from player sales: ", "losses in 5 of the last 7 years (–£173m in total), including a club-record –£104m in 2024/25, funded by transfer credit and a £124m loan."],
    ["Relegation (May 2026) ", "turns a structural deficit into a liquidity test. Summer-2026 sales of c. £187m buy time, not sustainability."],
  ];
  s.addText(
    bullets.flatMap(([b, t], i) => [
      { text: b, options: { bold: true, color: C.background1, bullet: true, paraSpaceAfter: 8 } },
      { text: t, options: { color: C.background2, breakLine: i < bullets.length - 1 } },
    ]),
    { x: 0.5, y: 3.08, w: 6.9, h: 2.75, fontSize: 14, valign: "top", margin: 0, isTextBox: true, objectName: "Answer bullets" }
  );
  card(s, 0.5, 5.95, 6.9, 0.62, C.accent2, "Verdict box");
  s.addText(
    [
      { text: "Verdict: ", options: { bold: true } },
      { text: "financially fragile – not sustainable without a cost and capital reset." },
    ],
    { x: 0.7, y: 5.95, w: 6.5, h: 0.62, fontSize: 15, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: "Verdict text" }
  );

  // right: scorecard
  card(s, 7.75, 2.65, 5.08, 3.92, C.accent1, "Scorecard card");
  s.addText("Financial health scorecard", { x: 8.0, y: 2.8, w: 4.6, h: 0.36, fontSize: 16, bold: true, color: C.background1, margin: 0, isTextBox: true, objectName: "Scorecard header" });
  const rows = [
    [fa.FaChartLine, "Revenue growth & mix", "Weak", C.accent4, C.background1],
    [fa.FaPoundSign, "Profitability", "Weak", C.accent4, C.background1],
    [fa.FaUsers, "Cost control (wages, amortisation)", "Weak", C.accent4, C.background1],
    [fa.FaBalanceScale, "Debt & liquidity", "Weak", C.accent4, C.background1],
    [fa.FaFutbol, "Assets: stadium lease, academy, brand", "Strong", C.accent6, C.background1],
    [fa.FaHandshake, "Owner support & governance", "Mixed", C.accent5, C.text1],
  ];
  for (let i = 0; i < rows.length; i++) {
    const [Ic, label, status, pill, pillText] = rows[i];
    const y = 3.3 + i * 0.53;
    s.addShape(pres.shapes.OVAL, { x: 8.0, y: y + 0.02, w: 0.4, h: 0.4, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Icon circle ${i + 1}` });
    s.addImage({ data: await icon(Ic, HEX.dk2), x: 8.1, y: y + 0.12, w: 0.2, h: 0.2, altText: label, objectName: `Icon ${i + 1}` });
    s.addText(label, { x: 8.55, y, w: 2.95, h: 0.44, fontSize: 13, color: C.background1, valign: "middle", margin: 0, isTextBox: true, objectName: `Score label ${i + 1}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 11.6, y: y + 0.06, w: 0.98, h: 0.32, rectRadius: 0.16, fill: { color: pill }, line: { type: "none" }, objectName: `Status pill ${i + 1}` });
    s.addText(status, { x: 11.6, y: y + 0.06, w: 0.98, h: 0.32, fontSize: 12, bold: true, color: pillText, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Status ${i + 1}` });
  }
  sourceLine(s, "Figures: WH Holding Ltd accounts (financial years to 31 May) as reported; full source list in the speaker notes of each slide.", true);
  s.addNotes(
    "KEY MESSAGE: West Ham is financially fragile. It built a cost base for a top-eight Premier League club on revenue that only reached that level in years with European football. " +
      "Revenue: record £269.7m in 2023/24, down 16% to £227.5m in 2024/25 (no Europe, 14th place). " +
      "Costs: wages £136.8m (2022/23) to £176m (2024/25), +29%; wages/revenue 77% vs a 65% Premier League average (Deloitte ARFF 2026). Player purchases c. £570m from summer 2022 to summer 2025. " +
      "Results: pre-tax losses in 5 of 7 years 2018/19–2024/25 (cumulative c. –£173m); record –£104.2m loss in 2024/25. The only big profit (+£57.2m, 2023/24) came from selling Declan Rice. " +
      "Relegation on 24 May 2026 cuts revenue by roughly 40%; summer 2026 sales (c. £187m in, c. £39m out) provide liquidity, not a sustainable model. " +
      "Scorecard: weak on revenue, profit, cost and debt; strong assets (62,500-seat London Stadium on a low rent, academy, brand); owner support mixed – Křetínský (1890 Holdings) moving to c. 43% after David Sullivan stepped down as joint-chair in June 2026. " +
      "Sources: WH Holding Ltd financial statements (whufc.com/en/financial-reports) as reported by Sky Sports, Swiss Ramble, Matchday Finance and The Analysis Series; Deloitte Annual Review of Football Finance 2026; ESPN/Sky Sports (ownership); Irish Times/Fox Sports (relegation)."
  );
}

// ================================================================ SLIDE 2 – revenue
async function slide2() {
  pres.addSection({ title: "Analysis" });
  const s = pres.addSlide({ masterName: "Content Light", sectionTitle: "Analysis" });
  s.addText("Revenue: growth rode on Europe – then reversed", { placeholder: "title" });
  s.addText("Turnover rose from £140m (Covid) to a record £270m with three European campaigns; without Europe it fell 16% to £228m.", { placeholder: "subtitle" });

  const years = ["2018/19", "2019/20", "2020/21", "2021/22", "2022/23", "2023/24", "2024/25"];
  const revenue = [191, 140, 193, 253, 237, 270, 228];
  const europe = [false, false, false, true, true, true, false];
  s.addChart(pres.charts.BAR, [{ name: "Turnover £m", labels: years, values: revenue }], {
    x: 0.5, y: 1.72, w: 7.7, h: 4.4, ...chartBase(),
    barDir: "col", barGapWidthPct: 45,
    title: "Turnover, £m (financial years to 31 May)",
    chartColors: europe.map((e) => (e ? HEX.accent1 : HEX.accent3)),
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"£"0"m"', dataLabelFontSize: 12, dataLabelFontBold: true,
    valAxisMinVal: 0, valAxisMaxVal: 300, showLegend: false,
    objectName: "Revenue chart",
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.62, y: 6.24, w: 0.16, h: 0.16, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Legend Europe" });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.62, y: 6.24, w: 0.16, h: 0.16, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "Legend no Europe" });
  s.addText("European campaign that season", { x: 0.86, y: 6.17, w: 2.7, h: 0.3, fontSize: 11, color: MUTED, margin: 0, valign: "middle", isTextBox: true, objectName: "Legend text 1" });
  s.addText("No European football (2019/20 also hit by Covid)", { x: 3.86, y: 6.17, w: 4.3, h: 0.3, fontSize: 11, color: MUTED, margin: 0, valign: "middle", isTextBox: true, objectName: "Legend text 2" });

  // right column: revenue mix table
  s.addText("Revenue mix, £m", { x: 8.55, y: 1.72, w: 4.28, h: 0.34, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "Mix header" });
  const hdr = (t, align = "right") => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, align } });
  const cell = (t, o = {}) => ({ text: t, options: { align: "right", ...o } });
  const mix = [
    [hdr("", "left"), hdr("2023/24"), hdr("2024/25"), hdr("Change")],
    [cell("Broadcasting", { align: "left" }), cell("167"), cell("132"), cell("–21%", { color: C.accent4, bold: true })],
    [cell("Matchday", { align: "left" }), cell("45"), cell("39"), cell("–13%", { color: C.accent4, bold: true })],
    [cell("Commercial", { align: "left" }), cell("58"), cell("56"), cell("–3%", { color: C.accent4, bold: true })],
    [cell("Total", { align: "left", bold: true }), cell("270", { bold: true }), cell("228", { bold: true }), cell("–16%", { color: C.accent4, bold: true })],
  ];
  s.addTable(mix, {
    x: 8.55, y: 2.1, w: 4.28, colW: [1.48, 0.93, 0.93, 0.94], rowH: 0.33, fontSize: 13, color: C.text1, margin: [0.03, 0.08, 0.03, 0.08], valign: "middle",
    border: { type: "solid", pt: 0.75, color: GRID }, fill: { color: C.background2 }, objectName: "Revenue mix table",
  });

  s.addText("Benchmarks", { x: 8.55, y: 3.92, w: 4.28, h: 0.34, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "Benchmark header" });
  const stats = [
    ["20th", "Deloitte Money League 2026 (17th a year earlier) – lowest since 2021"],
    ["58%", "of 2024/25 revenue from broadcasting: income tracks league position"],
    ["£56m", "commercial income vs. £70m at Aston Villa – modest for a 62,500-seat London club"],
    ["£340m", "average Premier League club revenue 2024/25 (£6.8bn ÷ 20) vs. West Ham's £228m"],
  ];
  stats.forEach(([n, t], i) => {
    const y = 4.3 + i * 0.57;
    s.addText(n, { x: 8.55, y, w: 1.15, h: 0.5, fontSize: 22, bold: true, fontFace: "+mj-lt", color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: `Benchmark stat ${i + 1}` });
    s.addText(t, { x: 9.75, y, w: 3.08, h: 0.5, fontSize: 12, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: `Benchmark text ${i + 1}` });
  });
  sourceLine(s, "Sources: WH Holding Ltd accounts as reported by club releases, Swiss Ramble, Matchday Finance, Sky Sports; Deloitte Football Money League 2026 and Annual Review of Football Finance 2026; SportsPro (Aston Villa).");
  s.addNotes(
    "KEY MESSAGE: Revenue growth was real but not structural – it came from European runs and Premier League merit payments, and it reversed as soon as those stopped. " +
      "Trend (WH Holding Ltd, years to 31 May): 2018/19 £190.7m; 2019/20 c. £140m (Covid, c. £26m deferred into 2020/21); 2020/21 £192.7m; 2021/22 c. £253m (Europa League semi-final); 2022/23 £236.7m (Conference League winners); 2023/24 £269.7m record (Europa League quarter-final); 2024/25 £227.5m (no Europe, 14th). " +
      "Mix 2024/25: broadcasting £132.4m (–£34.6m: lost UEFA distributions and lower merit payment, 9th to 14th); matchday c. £39m (c. £45m); commercial c. £56m (c. £58m). Broadcasting is 58% of revenue, so income is highly exposed to sporting results. " +
      "Benchmarks: Deloitte Money League 2026 rank 20th (revenue €276m / £240.4m on Deloitte's basis), down from 17th; Tottenham 9th, Aston Villa 14th, Newcastle 17th. Premier League clubs' aggregate revenue £6.8bn in 2024/25 (Deloitte), c. £340m per club. Aston Villa commercial income £70m (SportsPro). " +
      "Implication: even before relegation the club's revenue ceiling without Europe was c. £230m. Relegation now removes most Premier League TV money (c. £132m falls to roughly £60m including the parachute payment). " +
      "Note: some outlets report 2024/25 revenue as £226m–£232m and 2021/22 as £253m–£255m depending on club vs group basis."
  );
}

// ================================================================ SLIDE 3 – squad investment & spending
async function slide3() {
  const s = pres.addSlide({ masterName: "Content Light", sectionTitle: "Analysis" });
  s.addText("Squad spending outran income", { placeholder: "title" });
  s.addText("Wages rose 29% in two years while revenue fell 4%; the only recent profit came from selling academy graduate Declan Rice.", { placeholder: "subtitle" });

  const years = ["2018/19", "2019/20", "2020/21", "2021/22", "2022/23", "2023/24", "2024/25"];
  const revenue = [191, 140, 193, 253, 237, 270, 228];
  const wages = [136, 131, 129, 136, 137, 161, 176];
  s.addChart(
    pres.charts.BAR,
    [
      { name: "Revenue", labels: years, values: revenue },
      { name: "Wages", labels: years, values: wages },
    ],
    {
      x: 0.5, y: 1.72, w: 7.7, h: 3.05, ...chartBase(),
      barDir: "col", barGrouping: "clustered", barGapWidthPct: 55, barOverlapPct: -5,
      title: "Revenue vs. wages, £m – wages/revenue: 54% (2021/22) → 77% (2024/25)",
      chartColors: [HEX.accent3, HEX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0", dataLabelFontSize: 10, dataLabelFontBold: true,
      valAxisMinVal: 0, valAxisMaxVal: 300,
      showLegend: true, legendPos: "t", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: MUTED,
      objectName: "Revenue vs wages chart",
    }
  );

  s.addChart(pres.charts.BAR, [{ name: "Pre-tax result £m", labels: years, values: [-28, -65, -27, 12, -18, 57, -104] }], {
    x: 0.5, y: 4.85, w: 7.7, h: 1.78, ...chartBase(),
    barDir: "col", barGapWidthPct: 45,
    title: "Pre-tax result, £m – cumulative 2018/19–2024/25: –£173m",
    chartColors: [HEX.accent6], invertedColors: [HEX.accent4],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"+£"0"m";"–£"0"m"', dataLabelFontBold: true,
    valAxisMinVal: -140, valAxisMaxVal: 90, catAxisLabelPos: "low", showLegend: false,
    objectName: "Pre-tax chart",
  });

  // right column: four stat cards
  const cards = [
    ["c. £570m", "gross player purchases from summer 2022 to summer 2025, incl. a record £184m in 2022/23 – only the 2023 window (Rice) was net-positive"],
    ["£99m", "player amortisation in 2024/25 (+19%); wages plus amortisation reached £275m – 121% of revenue"],
    ["£96m to £20m", "profit on player sales, 2023/24 vs. 2024/25: without a Rice-type sale the club swung from +£57m to –£104m"],
    ["£192m + £124m", "net transfer debt (May 2025) plus a five-year loan from Rights & Media Funding (£89m drawn); c. £15m implied interest on deferred fees"],
  ];
  cards.forEach(([n, t], i) => {
    const y = 1.72 + i * 1.23;
    card(s, 8.55, y, 4.28, 1.13, C.background2, `Spend card ${i + 1}`);
    s.addText(n, { x: 8.75, y: y + 0.08, w: 3.9, h: 0.4, fontSize: 22, bold: true, fontFace: "+mj-lt", color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: `Spend stat ${i + 1}` });
    s.addText(t, { x: 8.75, y: y + 0.48, w: 3.9, h: 0.6, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: `Spend text ${i + 1}` });
  });
  sourceLine(s, "Sources: WH Holding Ltd accounts as reported (Swiss Ramble, Matchday Finance, Inside World Football, The Analysis Series); Sports Mole (summer-2025 window). Purchases = accounts additions 2022/23–2024/25 + summer-2025 fees.");
  s.addNotes(
    "KEY MESSAGE: The cost base grew on the assumption of permanent European football. When revenue fell, costs did not – profit came only from selling the best academy asset. " +
      "Wages (£m): 2018/19 136; 2019/20 131 (incl. Pellegrini severance; ratio 94%, highest in the Premier League that year); 2020/21 129; 2021/22 135.7; 2022/23 136.8; 2023/24 161 (+18%); 2024/25 176 (+9%). Ratio rose from 54% (2021/22) to 77% (2024/25). " +
      "Squad investment: record £183.9m in summer 2022 (Paquetá, Scamacca, Aguerd…); £124m player purchases in 2023/24; £133m in 2024/25 (Kilman, Summerville, Füllkrug, Wan-Bissaka…); £128.5m in summer 2025 (Mateus Fernandes £38.5m, Diouf, Hermansen…). Over five seasons to 2025/26, c. €700m spent vs c. €330m received (Squawka); only one net-positive window (Rice to Arsenal, 2023). " +
      "Amortisation £99m in 2024/25 (up £16m); operating loss c. £105m before player trading. Profit on player sales £96.3m (2023/24) vs £20.0m (2024/25). " +
      "Pre-tax (£m): –28.2, –65.3, –26.9, +12.3, –18.3, +57.2, –104.2 = c. –£173m over seven years. " +
      "Financing: net transfer-related debt £192m at May 2025 (c. £196m owed to other clubs in instalments); Barclays overdraft c. £16m; new £124m five-year term loan from Rights and Media Funding Ltd (July 2025, £89m drawn by Feb 2026). Interest payable c. £18–22m in 2024/25, of which £15.2m implied interest on transfer fees with extended payment terms. " +
      "Implication: the club financed squad ambition with deferred transfer payments and debt – a model that only works with continuous European/Premier League revenue."
  );
}

// ================================================================ SLIDE 4 – benchmarks & relegation stress test
async function slide4() {
  const s = pres.addSlide({ masterName: "Content Light", sectionTitle: "Analysis" });
  s.addText("Benchmarks and the relegation stress test", { placeholder: "title" });
  s.addText("West Ham already spent more of its income on wages than most peers; relegation removes c. 40% of revenue faster than costs can fall.", { placeholder: "subtitle" });

  // horizontal bar: wages / revenue 2024/25 (listed bottom-to-top)
  const clubs = ["Wolves", "Bournemouth", "West Ham", "Brentford", "PL average", "Liverpool", "Arsenal", "Tottenham"];
  const vals = [95, 80, 77, 76, 65, 61, 50, 45];
  const cols = clubs.map((c) => (c === "West Ham" ? HEX.accent1 : c === "PL average" ? MUTED : HEX.accent3));
  s.addChart(pres.charts.BAR, [{ name: "Wages / revenue 2024/25", labels: clubs, values: vals }], {
    x: 0.5, y: 1.72, w: 6.0, h: 4.2, ...chartBase(),
    barDir: "bar", barGapWidthPct: 40,
    title: "Wages as % of revenue, 2024/25",
    chartColors: cols, catAxisLabelFontSize: 12,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', dataLabelFontBold: true,
    valAxisMinVal: 0, valAxisMaxVal: 105, showLegend: false,
    objectName: "Wage ratio benchmark chart",
  });
  s.addText(
    "Wolves and West Ham – two of the three clubs relegated in 2025/26 – had among the highest ratios. New EFL Squad Cost Ratio cap from 2026/27: 85% of income (wider cost base incl. amortisation).",
    { x: 0.5, y: 5.98, w: 6.0, h: 0.62, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "Benchmark note" }
  );

  // right: stress test table
  s.addText("Relegation stress test, £m (2026/27 = illustrative estimate)", { x: 6.85, y: 1.72, w: 5.98, h: 0.34, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "Stress test header" });
  const hdr = (t, align = "right") => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, align } });
  const cell = (t, o = {}) => ({ text: t, options: { align: "right", ...o } });
  const rows = [
    [hdr("", "left"), hdr("2024/25 PL (actual)"), hdr("2026/27 Champ. (est.)")],
    [cell("Broadcasting", { align: "left" }), cell("132"), cell("c. 60 (parachute c. 55)")],
    [cell("Matchday", { align: "left" }), cell("39"), cell("c. 30")],
    [cell("Commercial", { align: "left" }), cell("56"), cell("c. 40")],
    [cell("Total revenue", { align: "left", bold: true }), cell("228", { bold: true }), cell("c. 130  (–43%)", { bold: true, color: C.accent4 })],
    [cell("Wages", { align: "left", bold: true }), cell("176", { bold: true }), cell("c. 90–110 (clauses)", { bold: true })],
  ];
  s.addTable(rows, {
    x: 6.85, y: 2.1, w: 5.98, colW: [1.78, 1.9, 2.3], rowH: 0.33, fontSize: 13, color: C.text1, margin: [0.03, 0.08, 0.03, 0.08], valign: "middle",
    border: { type: "solid", pt: 0.75, color: GRID }, fill: { color: C.background2 }, objectName: "Stress test table",
  });

  const box = (x, title, color, items, name) => {
    card(s, x, 4.3, 2.89, 2.3, C.background2, name + " card");
    s.addText(title, { x: x + 0.18, y: 4.4, w: 2.55, h: 0.32, fontSize: 14, bold: true, color, margin: 0, isTextBox: true, objectName: name + " header" });
    s.addText(
      items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 5 } })),
      { x: x + 0.18, y: 4.76, w: 2.6, h: 1.78, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: name + " bullets" }
    );
  };
  box(6.85, "Cushions", C.accent6, [
    "Summer 2026: c. £187m sales vs. £39m spent (Fernandes £85m to Spurs)",
    "Parachute: 55% / 45% / 20% of PL equal share over 3 years",
    "Relegation wage clauses; stadium rent halved",
  ], "Cushions");
  box(9.94, "Risks", C.accent4, [
    "Leicester 2023/24: Championship wages £107m > revenue £105m",
    "8th when Nuno was replaced by Scott Parker (Nov 2026)",
    "Loan and transfer instalments still due",
  ], "Risks");
  sourceLine(s, "Sources: Two-Footed Tackle, PSR Watch (club ratios); Deloitte ARFF 2026 (PL average); EFL (May 2026); Leonard Curtis (Sep 2026); GLA (rent); Sports Mole, WestHamWorld (2026 fees). 2026/27 column = own estimate.");
  s.addNotes(
    "KEY MESSAGE: Benchmarked against peers, West Ham's cost structure was already that of a club living beyond its means; relegation exposes it. " +
      "Wages/revenue 2024/25: Tottenham c. 45%, Arsenal 50%, Liverpool 61%, Premier League average 65% (Deloitte), Brentford 76% (£131m on £173m), West Ham 77%, Bournemouth 80%, Wolves 95%. West Ham pays a Brentford/Bournemouth-level share of revenue despite a much bigger stadium and brand – i.e. the problem is cost discipline, not size. " +
      "Championship context: wages/turnover across the Championship has not fallen below 91% since 2015/16 (Leonard Curtis, Sep 2026). From 2026/27 the EFL replaces P&S rules with a Squad Cost Ratio cap of 85% of income with real-time monitoring; breaches trigger a levy. " +
      "Stress test (own illustrative estimate): broadcasting falls from £132m to c. £60m (first-year parachute c. £55m = 55% of the PL equal share, plus EFL distributions); matchday c. £30m with Championship pricing; commercial c. £40m after relegation clauses in sponsorship deals; total c. £130m (–43%). Wages must fall from £176m to roughly £90–110m – relegation clauses can cut many contracts by up to 50%. " +
      "Cushions: summer 2026 sales c. £187m vs c. £39m purchases (Mateus Fernandes to Tottenham for a club-record £85m; Diouf, Summerville, Potts also left); parachute payments for three seasons (55%/45%/20%); London Stadium rent halved under the lease (c. £4.4–4.6m to c. £2.1–2.2m). Parachute clubs earn 2.8x the revenue of other Championship clubs and are four times more likely to be promoted (Leonard Curtis). " +
      "Risks: Leicester precedent – Championship wages £107.2m exceeded revenue £105.3m in 2023/24 and led to PSR trouble; sporting risk (8th when Nuno was sacked on 3 Nov 2026, Scott Parker appointed); debt service on the £124m loan and transfer instalments; proposals to reduce parachute payments in the new PL–EFL deal."
  );
}

// ================================================================ SLIDE 5 – recommendations
async function slide5() {
  pres.addSection({ title: "Recommendations" });
  const s = pres.addSlide({ masterName: "Content Dark", sectionTitle: "Recommendations" });
  s.addText("Recommendations: priorities to 2028/29", { placeholder: "title" });
  s.addText("Reset costs, de-lever and change the squad model – and use the parachute window to return on a sustainable footing.", { placeholder: "subtitle" });

  const recs = [
    ["Reset the cost base to Championship reality", "Enforce relegation wage clauses, sell or loan high earners, and give every new contract relegation/promotion clauses and a 3-year maximum term.", "Squad cost ratio ≤ 75% (10 pts under the EFL cap); wage bill c. £100m in 2026/27"],
    ["Recapitalise and de-lever", "Use the c. £150m net summer-2026 proceeds to clear transfer instalments and repay the £124m loan; new lead owner (Křetínský, c. 43%) to fund through equity, not debt.", "Net transfer debt < £75m and interest < £8m a year by 2028/29"],
    ["Switch to “buy young – develop – sell”", "Data-led sporting director, under-24 recruitment with resale value (Brighton/Brentford model), clear academy pathway – Rice alone produced £96m profit.", "≥ £30m profit on player sales a year; no transfer fee above 15% of revenue"],
    ["Protect and diversify revenue", "Keep the 62,500-seat stadium full with Championship pricing (rent is halved), add promotion-linked sponsor uplifts, grow international and women's football partnerships.", "Matchday ≥ £30m in the Championship; commercial ≥ £70m after promotion"],
    ["Win promotion inside the parachute window", "Aim for promotion in 2026/27 or 2027/28 while parachute payments are 55% / 45%; agree a plan B now (non-parachute budget) if not promoted by May 2028.", "Premier League by 2028/29 with wages ≤ 65% of revenue (PL average)"],
  ];
  recs.forEach(([head, detail, kpi], i) => {
    const y = 1.74 + i * 0.82;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.1, w: 0.5, h: 0.5, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Number circle ${i + 1}` });
    s.addText(String(i + 1), { x: 0.5, y: y + 0.1, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Number ${i + 1}` });
    s.addText(head, { x: 1.18, y, w: 7.45, h: 0.3, fontSize: 15, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: `Priority ${i + 1}` });
    s.addText(detail, { x: 1.18, y: y + 0.31, w: 7.45, h: 0.44, fontSize: 12, color: C.background2, margin: 0, valign: "top", isTextBox: true, objectName: `Priority detail ${i + 1}` });
    card(s, 8.86, y, 3.97, 0.74, C.accent1, `KPI card ${i + 1}`);
    s.addText(
      [
        { text: "TARGET", options: { bold: true, color: C.accent2, fontSize: 10, charSpacing: 1, breakLine: true } },
        { text: kpi, options: { color: C.background1, fontSize: 12 } },
      ],
      { x: 9.02, y: y + 0.04, w: 3.7, h: 0.66, margin: 0, valign: "middle", isTextBox: true, objectName: `KPI ${i + 1}` }
    );
  });

  const phases = [
    ["2026/27  STABILISE", "Cut wages and debt, comply with SCR, push for the top two"],
    ["2027/28  REBUILD", "New recruitment model in place; promotion with 45% parachute"],
    ["2028/29  SUSTAIN", "Premier League return; break-even before player trading"],
  ];
  phases.forEach(([h, t], i) => {
    const x = 0.5 + i * 4.18;
    s.addShape(pres.shapes.CHEVRON, { x, y: 5.9, w: 3.97, h: 0.72, fill: { color: i === 0 ? C.accent2 : C.background2 }, line: { type: "none" }, objectName: `Phase ${i + 1}` });
    s.addText(
      [
        { text: h, options: { bold: true, fontSize: 12, breakLine: true } },
        { text: t, options: { fontSize: 11 } },
      ],
      { x: x + 0.42, y: 5.9, w: 3.15, h: 0.72, color: C.text2, margin: 0, valign: "middle", isTextBox: true, objectName: `Phase text ${i + 1}` }
    );
  });
  sourceLine(s, "Targets are proposals derived from the analysis on slides 2–4; financial years end 31 May.", true);
  s.addNotes(
    "KEY MESSAGE: The priority is not to 'spend our way back' but to reset the model so that promotion leads to a sustainable Premier League club rather than a repeat of 2024/25. " +
      "1) Cost reset: relegation clauses (up to c. 50% wage cuts), exit high earners, cap contract length; manage to an internal squad cost ratio of 75% to keep 10 points of headroom under the EFL's 85% cap. " +
      "2) De-lever: net summer-2026 proceeds c. £148–150m should prioritise transfer creditors (c. £192m net) and the expensive £124m Rights & Media Funding loan; the new majority shareholder should provide equity rather than further secured debt; complete the ownership transition after David Sullivan stepped down (June 2026) with a stronger, more independent board and a published three-year financial plan – also preparing for the Independent Football Regulator's licensing regime. " +
      "3) Squad model: shift from expensive, older signings (c. £570m since 2022, low resale) to Brighton/Brentford-style trading of young assets; use the academy (Rice sale = £96m profit) as a recurring profit source. " +
      "4) Revenue: the 62,500-seat London Stadium is the club's biggest structural advantage – keep it full with affordable Championship prices; build sponsorship contracts with promotion uplifts; international partnerships and the women's team diversify income. " +
      "5) Promotion window: parachute payments are 55% of the PL equal share in year 1 and 45% in year 2 (20% in year 3); parachute clubs are four times more likely to be promoted, so the window is 2026/27–2027/28. A pre-agreed plan B avoids the Leicester-style overspend if promotion fails. " +
      "Timeline: 2026/27 stabilise; 2027/28 rebuild and promote; 2028/29 sustain in the Premier League with wages at or below the 65% league average and break-even before player trading."
  );
}

// pptxgenjs 4.0.1 writes an <a:pPr> before *every* run, so a paragraph that mixes
// bold and regular runs gets a second <a:pPr> after an <a:r> – invalid DrawingML.
// Keep only the leading <a:pPr> of each paragraph.
async function stripMidParagraphPPr(file) {
  const fs = require("fs");
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  let removed = 0;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    const fixed = xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) =>
      "<a:p>" + inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, (pp, offset) => (offset === 0 ? pp : (removed++, ""))) + "</a:p>"
    );
    zip.file(name, fixed);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return removed;
}

(async () => {
  await slide1();
  await slide2();
  await slide3();
  await slide4();
  await slide5();
  await pres.writeFile({ fileName: OUT });
  console.log("removed mid-paragraph pPr:", await stripMidParagraphPPr(OUT));
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
