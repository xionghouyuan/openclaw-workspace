const puppeteer = require('puppeteer-core');
const path = require('path');

const files = [
  ['/tmp/slide1.html', '/tmp/slide1.jpg'],
  ['/tmp/slide2.html', '/tmp/slide2.jpg'],
];

(async () => {
  const browser = await puppeteer.launch({
    executablePath: '/usr/bin/google-chrome',
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  for (const [htmlFile, outFile] of files) {
    const page = await browser.newPage();
    await page.setViewport({ width: 1600, height: 900, deviceScaleFactor: 2 });
    await page.goto('file://' + path.resolve(htmlFile), { waitUntil: 'networkidle0' });
    const bodyHeight = await page.evaluate(() => document.body.scrollHeight);
    await page.setViewport({ width: 1600, height: bodyHeight, deviceScaleFactor: 2 });
    await page.screenshot({ type: 'jpeg', quality: 92, path: outFile, fullPage: true });
    console.log('Saved:', outFile);
    await page.close();
  }

  await browser.close();
  console.log('All done');
})();
