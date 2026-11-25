// Builds the one-page Word version of the sunk-cost assignment.
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, LevelFormat, VerticalAlign } = require('docx');

const CLARET = '7A263A', SOFT = 'F6EDF0', BLUE = 'E9F4FA', LINE = 'D9CFD3';
const FONT = 'Arial';
const r = (t, o = {}) => new TextRun({ text: t, font: FONT, ...o });
// Parse **bold** segments
const runs = (s, o = {}) => s.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(p =>
  p.startsWith('**') ? r(p.slice(2, -2), { ...o, bold: true }) : r(p, o));
const para = (s, o = {}, po = {}) => new Paragraph({ children: runs(s, o), spacing: { after: 40 }, ...po });
const bullet = (s) => new Paragraph({ children: runs(s), numbering: { reference: 'b', level: 0 }, spacing: { after: 30 } });

const W = [1500, 3050, 2500, 3006]; // sums to 10056 (A4 width minus 0.72in margins)
const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };
const cell = (children, i, fill) => new TableCell({
  children, width: { size: W[i], type: WidthType.DXA }, borders,
  shading: fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined,
  margins: { top: 60, bottom: 60, left: 90, right: 90 }, verticalAlign: VerticalAlign.TOP,
});
const head = (t, i) => cell([new Paragraph({ children: [r(t, { bold: true, color: 'FFFFFF' })] })], i, CLARET);

const areas = [
  {
    name: '1 Players', sub: 'transfer fees',
    costs: ['Transfer, agent and signing-on fees, plus scouting and relocation. Future wages are **not** sunk.',
      'Haller £45m (sold for ~£20m), Paquetá £51m (sold for ~£36.5m), ~£130m on the 2024 rebuild, Castellanos ~€29m in Jan 2026.'],
    bias: ['“He cost €29m, he has to start.”', '“We can’t sell below what we paid.” Castellanos wants out after relegation, and the club reportedly won’t sell “on the cheap”.'],
    recs: ['Select with a dashboard that leaves out fees and wages.', 'Each window, ask: “Would we sign him today?”', 'Set asking prices from market valuations, not the purchase price.'],
  },
  {
    name: '2 Head coach', sub: 'hire and fire',
    costs: ['Pay-offs (Lopetegui ~£3–5m, Potter ~£5m), staff compensation, coach-specific signings, lost time.',
      'Four coaches appointed since May 2024: Lopetegui, Potter, Nuno (relegated May 2026, sacked Nov 2026), Parker.'],
    bias: ['“We backed him with £100m+, so give him time.”', '“Sacking him admits our mistake.” A new coach may also discard useful players as “not my signing”.'],
    recs: ['A playing model owned by the club, not by the coach.', 'Forward-looking KPIs reviewed every 10 matches.', 'A review panel separate from the people who hired him, plus break clauses.'],
  },
  {
    name: '3 Stadium and seats', sub: 'London Stadium, season tickets',
    costs: ['Boleyn Ground left (2016), £15m conversion contribution, ~£750m public money (not ours).',
      '2025/26: 62,455 tickets sold per match, only 49,567 in the stadium. Relegation cut rent to ~£2.1m; the Mayor is open to a sale (~£50m reported).'],
    bias: ['“We’ve invested so much, so we must buy it.”', 'Fans: “I’ve paid, so I’ll go.” This fades during the season, so “sold out” still leaves ~12,900 empty seats.'],
    recs: ['Base buy-or-rent on a future cash-flow case only.', 'Offer monthly season-ticket payments.', 'Use a ticket exchange with a “use it or lose it” rule, and turnstile occupancy as the KPI.'],
  },
];

const rows = [new TableRow({ tableHeader: true, children: [
  head('Area', 0), head('Sunk costs and how they accrued', 1), head('Where the bias shows up', 2), head('Recommendations', 3)] })];
areas.forEach((a, k) => rows.push(new TableRow({ children: [
  cell([new Paragraph({ children: [r(a.name, { bold: true, color: CLARET })] }), para(a.sub, { color: '5D5D68', size: 19 })], 0, SOFT),
  cell(a.costs.map(bullet), 1),
  cell(a.bias.map(s => new Paragraph({ children: runs(s, { italics: true }), spacing: { after: 40 } })), 2, SOFT),
  cell(a.recs.map(bullet), 3, BLUE),
] })));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 220, hanging: 180 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 900, left: 925, right: 925 } } },
    children: [
      new Paragraph({ children: [r('Sunk Costs at West Ham United: Three Areas', { bold: true, size: 32, color: CLARET })], spacing: { after: 40 } }),
      new Paragraph({ children: [r('International Sport Management Game · West Ham United · November 2026 · Prepared by: [Team names]', { size: 18, color: '5D5D68' })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: CLARET, space: 4 } }, spacing: { after: 160 } }),
      para('A **sunk cost** is money already spent that cannot be recovered, whatever we decide next. Rationally, only **future** costs and benefits should count. In practice, decision-makers often let past spending steer new decisions: 54% chose the more expensive ski trip they would enjoy less, and 85% would finish the radar-invisible plane (Arkes & Blumer, 1985). Sports studies mostly confirm this (Staw & Hoang, 1995; Keefer, 2017; Hackinger, 2019), and Özaydın (2022) found the bias among Premier League coaches. We identify three areas where it threatens West Ham’s decisions after relegation in May 2026.', {}, { spacing: { after: 160 } }),
      new Table({ width: { size: 10056, type: WidthType.DXA }, columnWidths: W, rows }),
      new Paragraph({ spacing: { before: 180, after: 60 }, children: runs('**Conclusion.** In all three areas the risk is the same: past spending is used to justify the next decision. West Ham should make every major decision with forward-looking numbers. Each decision needs clear owners, review points agreed in advance, and someone who did not make the original decision. **Don’t let yesterday’s investment determine tomorrow’s decision.**') }),
      new Paragraph({ spacing: { before: 120 }, children: [r('Sources: Arkes & Blumer (1985), OBHDP; Staw & Hoang (1995), ASQ; Keefer (2017), J. Sports Econ.; Hackinger (2019), J. Econ. Psych.; Özaydın (2022), Pamukkale J. Sport Sci.; Staw, Barsade & Koput (1997), J. Appl. Psych.; Gourville & Soman (1998), J. Consumer Res.; Sky Sports, ESPN, Football Italia, Sports Mole, FourFourTwo (FOI turnstile data), GLA, City AM (2024–2026). Fees are as reported in the media.', { size: 16, color: '5D5D68' })] }),
    ],
  }],
});

const out = path.resolve(__dirname, '..', 'WestHam_SunkCosts_OnePage.docx');
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log('wrote', out); });
