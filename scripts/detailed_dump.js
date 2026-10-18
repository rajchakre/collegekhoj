// scripts/detailed_dump.js
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');

const DELIVERABLES_DIR = 'C:\\Users\\shlok\\Downloads\\Web_Development__Deliverables\\Deliverables';

async function dumpExcel() {
  const xlsxFiles = fs.readdirSync(DELIVERABLES_DIR).filter(f => f.endsWith('.xlsx'));
  for (const f of xlsxFiles) {
    console.log(`\n===============================================================`);
    console.log(`EXCEL: ${f}`);
    console.log(`===============================================================`);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(DELIVERABLES_DIR, f));
    for (const ws of wb.worksheets) {
      console.log(`\n--- Sheet: "${ws.name}" ---`);
      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const rowVals = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          let v = cell.value;
          if (v && typeof v === 'object') {
            if (v.richText) v = v.richText.map(t => t.text).join('');
            else if (v.result !== undefined) v = `[formula result: ${v.result}]`;
            else if (v.text !== undefined) v = v.text;
            else v = JSON.stringify(v);
          }
          if (v !== undefined && v !== null && String(v).trim() !== '') {
            rowVals.push(`Col ${colNumber}: "${String(v).trim()}"`);
          }
        });
        if (rowVals.length) {
          console.log(`Row ${rowNumber}: ${rowVals.join(' | ')}`);
        }
      });
    }
  }
}

dumpExcel().catch(e => console.error(e));
