const puppeteer = require('puppeteer');
const fs = require('fs-extra');
const mammoth = require('mammoth');
const path = require('path');
const config = require('./config');

async function convertToPDF() {
  console.log('[PDF] Converting document to PDF format...');

  try {
    await fs.ensureDir(path.dirname(config.PDF_PATH));

    // Extract text content from DOCX
    const result = await mammoth.extractRawText({ path: config.DOCX_PATH });
    
    // Create HTML wrapper for browser rendering
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="UTF-8">
          <title>${config.TITLE}</title>
          <style>
            body { 
              font-family: Arial, sans-serif; 
              font-size: 14pt; 
              line-height: 1.5;
              padding: 2cm;
              margin: 0;
            }
            h1 { 
              color: ${config.COLORS.primary}; 
              border-bottom: 2px solid ${config.COLORS.gold}; 
              padding-bottom: 10px; 
            }
            h2 { 
              color: ${config.COLORS.primary}; 
            }
            h3 { 
              color: ${config.COLORS.secondary}; 
            }
            table { 
              width: 100%; 
              border-collapse: collapse; 
              margin: 20px 0; 
            }
            th, td { 
              border: 1px solid #ccc; 
              padding: 8px; 
              text-align: left; 
            }
            th { 
              background-color: ${config.COLORS.primary}; 
              color: white; 
            }
            tr:nth-child(even) { 
              background-color: ${config.COLORS.lightGray}; 
            }
            .page-break { 
              page-break-before: always; 
            }
            .center { 
              text-align: center; 
            }
            .bold { 
              font-weight: bold; 
            }
            .italic { 
              font-style: italic; 
            }
          </style>
        </head>
        <body>${result.value.replace(/\n/g, '<br>')}</body>
      </html>
    `;

    const htmlPath = path.join(__dirname, '../temp.html');
    await fs.writeFile(htmlPath, htmlContent);

    // Launch headless browser and convert to PDF
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    await page.setContent(await fs.readFile(htmlPath, 'utf8'));

    await page.pdf({
      path: config.PDF_PATH,
      format: 'A4',
      printBackground: true,
      margin: { top: '2cm', right: '2cm', bottom: '2cm', left: '2cm' }
    });

    await browser.close();
    await fs.remove(htmlPath);

    console.log(`[SUCCESS] PDF created: ${config.PDF_PATH}`);
    console.log(`[FILE SIZE] ${formatSize(await fs.stat(config.PDF_PATH))}`);

  } catch (error) {
    console.error('[ERROR] PDF conversion failed:', error.message);
    throw error;
  }
}

function formatSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = { convertToPDF };
if (require.main === module) {
  convertToPDF().catch(err => {
    console.error('Conversion failed:', err);
    process.exit(1);
  });
}
