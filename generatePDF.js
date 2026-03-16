const { execFile } = require('child_process');
const path = require('path');
const fs   = require('fs');

/**
 * Generates a PDF quote file.
 * Returns the path to the generated .pdf file.
 */
function generatePDF(quoteData, advisorName) {
    return new Promise((resolve, reject) => {
        const quotesDir = path.join(__dirname, 'quotes');
        if (!fs.existsSync(quotesDir)) fs.mkdirSync(quotesDir, { recursive: true });

        const quoteId    = quoteData.quote_id || 'UNKNOWN';
        const outputPath = path.join(quotesDir, `${quoteId}.pdf`);
        const logoFolder = path.join(__dirname, 'logos');
        const scriptPath = path.join(__dirname, 'generate_quote_pdf.py');

        // Add advisor name into data
        const data = { ...quoteData, advisor_name: advisorName || 'your trusted advisor' };
        const jsonArg = JSON.stringify(data);

        // Try python3 first, fall back to python
        const tryPython = (cmd) => new Promise((res, rej) => {
            execFile(cmd, [scriptPath, jsonArg, outputPath, logoFolder], 
                { maxBuffer: 10 * 1024 * 1024 },
                (err, stdout, stderr) => {
                    if (err) return rej(err);
                    res(stdout.trim() || outputPath);
                }
            );
        });

        tryPython('python3')
            .catch(() => tryPython('python'))
            .then(outPath => resolve(outPath || outputPath))
            .catch(err => reject(new Error(`Python script failed: ${err.message}`)));
    });
}

module.exports = generatePDF;
