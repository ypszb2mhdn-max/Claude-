// Renders the HTML sources in this folder to PDF with Chromium.
// Usage: node build.js <outDir>
const { chromium } = require('playwright');
const path = require('path');

const outDir = path.resolve(process.argv[2] || path.join(__dirname, '..'));

const only = process.argv[3];
const docs = [
  { src: 'analysis.html', out: 'WestHam_x_Pistons_Partnership_Analysis.pdf', footer: 'West Ham United × Detroit Pistons · Partnership Analysis' },
  { src: 'video-script.html', out: 'WestHam_x_Pistons_Video_Script.pdf', footer: 'West Ham United × Detroit Pistons · Video Script' },
  { src: 'ethics-assignment.html', out: 'WestHam_Ethics_Governance_Leadership.pdf', footer: 'West Ham United · Ethics, Governance and Leadership' },
];

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  for (const d of docs.filter(x => !only || x.src === only)) {
    await page.goto('file://' + path.join(__dirname, d.src), { waitUntil: 'networkidle' });
    await page.evaluate(() => document.fonts.ready);
    await page.pdf({
      path: path.join(outDir, d.out),
      format: 'A4',
      preferCSSPageSize: true,
      printBackground: true,
      displayHeaderFooter: true,
      headerTemplate: '<span></span>',
      footerTemplate: `<div style="width:100%;font-family:Inter,sans-serif;font-size:7.5px;color:#777;padding:0 17mm;display:flex;justify-content:space-between;"><span>${d.footer}</span><span><span class="pageNumber"></span> / <span class="totalPages"></span></span></div>`,
    });
    console.log('wrote', d.out);
  }
  await browser.close();
})();
