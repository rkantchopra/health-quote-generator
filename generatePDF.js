'use strict';
const puppeteer           = require('puppeteer');
const { buildQuoteHTML }  = require('./templates/quoteHTML');
const path                = require('path');
const fs                  = require('fs');

/**
 * Generates a PDF quote using Puppeteer (headless Chrome → pixel-perfect HTML→PDF).
 * Returns the path to the generated .pdf file.
 */
function generatePDF(quoteData, advisorName) {
  return new Promise(async (resolve, reject) => {
    const quotesDir = path.join(__dirname, 'quotes');
    if (!fs.existsSync(quotesDir)) fs.mkdirSync(quotesDir, { recursive: true });

    const quoteId    = quoteData.quote_id || 'UNKNOWN';
    const outputPath = path.join(quotesDir, `${quoteId}.pdf`);
    const data       = { ...quoteData, advisor_name: advisorName || 'your trusted advisor' };

    let browser;
    try {
      const html = buildQuoteHTML(data);

      browser = await puppeteer.launch({
        headless: 'new',
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--disable-gpu',
        ],
      });

      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: 'domcontentloaded' });
      await page.pdf({
        path: outputPath,
        format: 'A4',
        landscape: true,
        margin: { top: '0.7cm', bottom: '0.7cm', left: '0.7cm', right: '0.7cm' },
        printBackground: true,
      });

      resolve(outputPath);
    } catch (err) {
      reject(new Error(`PDF generation failed: ${err.message}`));
    } finally {
      if (browser) await browser.close();
    }
  });
}

module.exports = generatePDF;
