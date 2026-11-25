// Builds the 5-slide deck "West Ham United FC – Financial performance & sustainability".
// Run: NODE_PATH=<dir with pptxgenjs, react, react-dom, react-icons, sharp> node build_deck.js <out.pptx> <pptx-skill-dir>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || path.join(__dirname, "..", "WestHam_Financial_Sustainability_Short.pptx");
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

// ================================================================ SHORT VERSION – PAGE 1: analysis
async function page1() {
  pres.addSection({ title: "Analysis" });
  const s = pres.addSlide({ masterName: "Content Light", sectionTitle: "Analysis" });
  s.addText("West Ham: a Premier League cost base without the income", { placeholder: "title" });
  s.addText("Revenue only grew with European football, while wages and transfer spending kept rising – profit came only from selling players.", { placeholder: "subtitle" });

  const years = ["2018/19", "2019/20", "2020/21", "2021/22", "2022/23", "2023/24", "2024/25"];
  s.addChart(
    pres.charts.BAR,
    [
      { name: "Revenue", labels: years, values: [191, 140, 193, 253, 237, 270, 228] },
      { name: "Wages", labels: years, values: [136, 131, 129, 136, 137, 161, 176] },
    ],
    {
      x: 0.5, y: 1.72, w: 7.5, h: 3.4, ...chartBase(),
      barDir: "col", barGrouping: "clustered", barGapWidthPct: 55, barOverlapPct: -5,
      title: "Revenue vs. wages, £m (years to 31 May) – Europe in 2021/22 to 2023/24",
      chartColors: [HEX.accent3, HEX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0", dataLabelFontSize: 10, dataLabelFontBold: true,
      valAxisMinVal: 0, valAxisMaxVal: 300,
      showLegend: true, legendPos: "t", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: MUTED,
      objectName: "Revenue vs wages chart",
    }
  );
  s.addChart(pres.charts.BAR, [{ name: "Pre-tax result £m", labels: years, values: [-28, -65, -27, 12, -18, 57, -104] }], {
    x: 0.5, y: 5.2, w: 7.5, h: 1.42, ...chartBase(),
    barDir: "col", barGapWidthPct: 45,
    title: "Pre-tax result, £m (total: –£173m)",
    chartColors: [HEX.accent6], invertedColors: [HEX.accent4],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"+"0;"–"0', dataLabelFontBold: true,
    valAxisMinVal: -150, valAxisMaxVal: 100, catAxisLabelPos: "low", showLegend: false,
    objectName: "Pre-tax chart",
  });

  const cards = [
    ["£228m", "revenue 2024/25 (–16%): record £270m in 2023/24 depended on Europe"],
    ["77%", "wages/revenue 2024/25 (£176m) vs. 65% Premier League average"],
    ["c. £570m", "spent on players since 2022; £99m a year amortisation"],
    ["–£104m", "record loss 2024/25; only profit came from the Rice sale (2023/24)"],
    ["£316m", "debt: £192m owed for transfers + £124m new loan"],
  ];
  cards.forEach(([n, t], i) => {
    const y = 1.72 + i * 0.98;
    card(s, 8.35, y, 4.48, 0.88, C.background2, `Card ${i + 1}`);
    s.addText(n, { x: 8.5, y, w: 1.75, h: 0.88, fontSize: 20, bold: true, fontFace: "+mj-lt", color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: `Stat ${i + 1}` });
    s.addText(t, { x: 10.3, y, w: 2.42, h: 0.88, fontSize: 12, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: `Stat text ${i + 1}` });
  });
  sourceLine(s, "Sources: WH Holding Ltd accounts as reported (Sky Sports, Swiss Ramble, Matchday Finance, Inside World Football, The Analysis Series); Deloitte Annual Review of Football Finance 2026.");
  s.addNotes(
    "Revenue rose to a record £269.7m in 2023/24 thanks to three European campaigns, then fell 16% to £227.5m in 2024/25 (TV £167m to £132m). " +
      "Wages rose from £136.8m (2022/23) to £176m (2024/25), 77% of revenue vs a 65% league average. Player purchases c. £570m from summer 2022 to summer 2025; amortisation £99m. " +
      "Pre-tax results 2018/19–2024/25: –28, –65, –27, +12, –18, +57 (Rice sale, £96m profit on player sales), –104 = c. –£173m. Net transfer debt £192m (May 2025) plus a £124m loan from Rights & Media Funding (July 2025)."
  );
}

// ================================================================ SHORT VERSION – PAGE 2: verdict & recommendations
async function page2() {
  pres.addSection({ title: "Recommendations" });
  const s = pres.addSlide({ masterName: "Content Dark", sectionTitle: "Recommendations" });
  s.addText("Verdict and priorities to 2028/29", { placeholder: "title" });
  s.addText("Not sustainable in its current form – relegation (May 2026) cuts revenue by c. 40% faster than costs can fall.", { placeholder: "subtitle" });

  // left: relegation impact
  card(s, 0.5, 1.72, 4.6, 4.9, C.accent1, "Outlook card");
  s.addText("Championship 2026/27 (estimate)", { x: 0.72, y: 1.85, w: 4.2, h: 0.35, fontSize: 15, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "Outlook header" });
  const rows = [["Revenue", "£228m → c. £130m"], ["of which TV", "£132m → c. £60m"], ["Wages", "£176m → c. £90–110m"]];
  rows.forEach(([a, b], i) => {
    const y = 2.3 + i * 0.5;
    s.addText(a, { x: 0.72, y, w: 1.6, h: 0.42, fontSize: 14, color: C.background2, margin: 0, valign: "middle", isTextBox: true, objectName: `Outlook label ${i + 1}` });
    s.addText(b, { x: 2.3, y, w: 2.65, h: 0.42, fontSize: 15, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: `Outlook value ${i + 1}` });
  });
  s.addText(
    [
      { text: "Helps: c. £187m player sales in summer 2026, parachute payments (55% / 45% / 20%), relegation wage clauses, stadium rent halved.", options: { bullet: true, breakLine: true, paraSpaceAfter: 6 } },
      { text: "Hurts: £124m loan and transfer instalments, new EFL cap (squad costs ≤ 85% of income), 8th place when the manager changed.", options: { bullet: true } },
    ],
    { x: 0.72, y: 3.95, w: 4.2, h: 2.55, fontSize: 13, color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: "Outlook bullets" }
  );

  // right: recommendations
  const recs = [
    ["Cut costs to Championship level", "Use relegation clauses; short contracts. Target: squad costs ≤ 75% of income."],
    ["Reduce debt with owner equity", "Use sale money to pay transfer debt and the loan. Target: transfer debt < £75m by 2029."],
    ["Buy young, develop, sell", "Brighton/Brentford model and academy (Rice = £96m profit). Target: ≥ £30m sales profit a year."],
    ["Protect revenue", "Full 62,500 stadium with fair prices; sponsor deals with promotion bonuses."],
    ["Promotion within the parachute years", "Up by 2027/28, back in the PL by 2028/29 with wages ≤ 65% of revenue."],
  ];
  recs.forEach(([h, d], i) => {
    const y = 1.72 + i * 0.99;
    s.addShape(pres.shapes.OVAL, { x: 5.45, y: y + 0.08, w: 0.5, h: 0.5, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Number circle ${i + 1}` });
    s.addText(String(i + 1), { x: 5.45, y: y + 0.08, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Number ${i + 1}` });
    s.addText(h, { x: 6.15, y, w: 6.68, h: 0.34, fontSize: 16, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: `Priority ${i + 1}` });
    s.addText(d, { x: 6.15, y: y + 0.36, w: 6.68, h: 0.5, fontSize: 13, color: C.background2, margin: 0, valign: "top", isTextBox: true, objectName: `Priority detail ${i + 1}` });
  });
  sourceLine(s, "Sources: Deloitte ARFF 2026; EFL (May 2026); Sports Mole (2026 transfers); GLA (stadium rent). 2026/27 figures are an own estimate.", true);
  s.addNotes(
    "Relegation on 24 May 2026: TV income falls from £132m to roughly £60m including the first-year parachute payment (c. £55m). Estimated revenue c. £130m. " +
      "Cushions: summer-2026 sales c. £187m vs c. £39m spent (Mateus Fernandes £85m to Tottenham), parachute payments for three years, wage relegation clauses, London Stadium rent halved. " +
      "Recommendations: 1) cost reset, 2) de-lever using sale proceeds and owner equity (Křetínský c. 43%), 3) buy-develop-sell squad model, 4) protect matchday and commercial revenue, 5) promotion while parachute payments are highest."
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
  await page1();
  await page2();
  await pres.writeFile({ fileName: OUT });
  await stripMidParagraphPPr(OUT);
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
