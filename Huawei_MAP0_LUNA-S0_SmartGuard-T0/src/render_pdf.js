// Rendert src/anschlussplan.html nach PDF (A4 quer) und optional PNG-Vorschau je Blatt.
// Aufruf: node src/render_pdf.js [vorschau-verzeichnis]
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const src = path.resolve(__dirname, 'anschlussplan.html');
  const out = path.resolve(__dirname, '..', 'Anschlussplan_Huawei_MAP0_2xLUNA-S0_SmartGuard-T0.pdf');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('file://' + src);
  await page.pdf({ path: out, preferCSSPageSize: true, printBackground: true });
  if (process.argv[2]) {
    await page.setViewportSize({ width: 1123, height: 794 });
    await page.emulateMedia({ media: 'print' });
    const pages = await page.$$('.page');
    for (let i = 0; i < pages.length; i++) {
      await pages[i].screenshot({ path: path.join(process.argv[2], `blatt${i + 1}.png`) });
    }
  }
  await browser.close();
  console.log('PDF:', out);
})();
