// Renders the three infographics to PNG (2x) and single-page PDFs, then merges the PDFs.
// Usage: node build.js
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const path = require('path');

const outDir = path.resolve(__dirname, '..');
const docs = [
  { src: 'area1-players.html', out: 'WestHam_SunkCosts_1_Players' },
  { src: 'area2-coach.html', out: 'WestHam_SunkCosts_2_HeadCoach' },
  { src: 'area3-stadium.html', out: 'WestHam_SunkCosts_3_Stadium_Seats' },
];

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1000, height: 1200 }, deviceScaleFactor: 2 });
  const pdfs = [];
  for (const d of docs) {
    await page.goto('file://' + path.join(__dirname, d.src), { waitUntil: 'networkidle' });
    await page.evaluate(() => document.fonts.ready);
    const h = await page.evaluate(() => Math.ceil(document.querySelector('.poster').getBoundingClientRect().height));
    await page.screenshot({ path: path.join(outDir, d.out + '.png'), fullPage: true });
    const pdf = path.join(outDir, d.out + '.pdf');
    await page.pdf({ path: pdf, width: '1000px', height: h + 'px', printBackground: true, pageRanges: '1',
      margin: { top: '0', right: '0', bottom: '0', left: '0' } });
    pdfs.push(pdf);
    console.log('wrote', d.out, '(height ' + h + 'px)');
  }
  await browser.close();
  execFileSync('pdfunite', [...pdfs, path.join(outDir, 'WestHam_SunkCosts_All_3_Infographics.pdf')]);
  console.log('wrote combined PDF');
})();
