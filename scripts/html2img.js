const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const htmlFile = process.argv[2] || '/home/xionghouyuan2/workplan.html';
const outputFile = process.argv[3] || '/home/xionghouyuan2/workplan_16x9.jpg';

(async () => {
  const browser = await puppeteer.launch({
    executablePath: '/usr/bin/google-chrome',
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  
  const resolvedPath = path.resolve(htmlFile);
  if (!fs.existsSync(resolvedPath)) {
    console.error('File not found:', resolvedPath);
    await browser.close();
    process.exit(1);
  }

  await page.goto('file://' + resolvedPath, { waitUntil: 'networkidle0' });
  
  // Get the actual page height
  const bodyHeight = await page.evaluate(() => document.body.scrollHeight);
  
  // 16:9 = 1600x900
  await page.setViewport({ width: 1600, height: bodyHeight, deviceScaleFactor: 2 });
  
  await page.screenshot({
    type: 'jpeg',
    quality: 92,
    path: outputFile,
    fullPage: true
  });

  console.log('Saved:', outputFile);
  await browser.close();
})();
