const fs = require('fs');
const path = require('path');
const pdfjsLib = require('pdfjs-dist');
const ExcelJS = require('exceljs');

const regex = /(?<date>\d{2}-[A-Za-z]{3}-\d{4})\s+(?<description>.+?)\s+(?<amount>\d{1,3}(?:,\d{2,3})*\.\d{2})\s+(?<balance>\d{1,3}(?:,\d{2,3})*\.\d{2})(?<type>Dr|Cr)/g;

const PDF_FILE_PATH = path.join(__dirname, 'file.pdf');
const EXCEL_FILE_PATH = path.join(__dirname, 'parsed_transactions.xlsx');

async function logTextFromPdf(pdfPath) {
  try {
    if (!fs.existsSync(pdfPath)) {
      throw new Error(`File not found: ${pdfPath}`);
    }

    const data = new Uint8Array(fs.readFileSync(pdfPath));
    const pdfDocument = await pdfjsLib.getDocument({ data }).promise;
    const numPages = pdfDocument.numPages;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Transactions');

    worksheet.columns = [
      { header: 'Date', key: 'date', width: 15 },
      { header: 'Description', key: 'description', width: 50 },
      { header: 'Amount', key: 'amount', width: 15 },
      { header: 'Balance', key: 'balance', width: 15 },
      { header: 'Type', key: 'type', width: 10 }
    ];

    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const page = await pdfDocument.getPage(pageNum);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(item => item.str).join(' ');

      const matches = [...pageText.matchAll(regex)];
      matches.forEach((match, index) => {
        const { date, description, amount, balance, type } = match.groups;

        console.log(`\nEntry ${index + 1}:`);
        console.log("Date:", date);
        console.log("Description:", description);
        console.log("Amount:", amount);
        console.log("Balance:", balance);
        console.log("Type:", type);

        worksheet.addRow({
          date,
          description,
          amount: parseFloat(amount.replace(/,/g, '')),
          balance: parseFloat(balance.replace(/,/g, '')),
          type
        });
      });
    }

    await workbook.xlsx.writeFile(EXCEL_FILE_PATH);
    console.log(`\n Excel file saved to ${EXCEL_FILE_PATH}`);

  } catch (error) {
    console.error(`Error: ${error.message}`);
  }
}

logTextFromPdf(PDF_FILE_PATH);
