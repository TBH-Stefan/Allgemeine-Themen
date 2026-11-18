// Rendert src/anschlussplan.html nach PDF (A4 quer) und optional PNG-Vorschau.
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const src = path.resolve(__dirname, 'anschlussplan.html');
  const out = path.resolve(__dirname, '..', 'Anschlussplan_PV-Ueberschuss_SE8K_WOLF-CHA10.pdf');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('file://' + src);
  await page.pdf({ path: out, preferCSSPageSize: true, printBackground: true });
  if (process.argv[2]) {
    await page.setViewportSize({ width: 1600, height: 1100 });
    await page.emulateMedia({ media: 'print' });
    await page.screenshot({ path: process.argv[2], fullPage: true });
  }
  await browser.close();
  console.log('PDF:', out);
})();
