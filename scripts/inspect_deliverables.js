// scripts/inspect_deliverables.js
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');

const DELIVERABLES_DIR = 'C:\\Users\\shlok\\Downloads\\Web_Development__Deliverables\\Deliverables';

async function inspectAll() {
  const files = fs.readdirSync(DELIVERABLES_DIR);
  console.log('Total files found:', files.length);

  for (const file of files) {
    const fullPath = path.join(DELIVERABLES_DIR, file);
    const ext = path.extname(file).toLowerCase();
    console.log(`\n==================================================`);
    console.log(`FILE: ${file}`);
    console.log(`==================================================`);

    if (ext === '.xlsx') {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fullPath);
      console.log(`Sheets (${workbook.worksheets.length}):`);
      for (const sheet of workbook.worksheets) {
        console.log(`\n  Sheet: "${sheet.name}" (Rows: ${sheet.rowCount}, Cols: ${sheet.columnCount})`);
        let count = 0;
        sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (count < 10) {
            const values = [];
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              if (colNumber <= 10) {
                let val = cell.value;
                if (val && typeof val === 'object') {
                  val = val.result || val.text || JSON.stringify(val);
                }
                values.push(`C${colNumber}: ${String(val || '').trim()}`);
              }
            });
            if (values.length) {
              console.log(`    R${rowNumber}:`, values.join(' | '));
              count++;
            }
          }
        });
      }
    } else if (ext === '.docx') {
      const zip = new AdmZip(fullPath);
      const docXmlEntry = zip.getEntry('word/document.xml');
      if (docXmlEntry) {
        const xml = zip.readAsText(docXmlEntry);
        // Find placeholders or text patterns like <...>, [...], {PROJECT...}
        const matches = xml.match(/<w:t[^>]*>([^<]+)<\/w:t>/g) || [];
        const textRuns = matches.map(m => m.replace(/<[^>]+>/g, '')).join(' ');
        console.log('Preview text (first 500 chars):');
        console.log(textRuns.replace(/\s+/g, ' ').substring(0, 500));
        // Search for bracketed placeholders like [Project Name] or <Project>
        const placeholders = textRuns.match(/\[[^\]]+\]|<[^>]+>|\bTBD\b|\bTODO\b/g) || [];
        console.log('Detected bracketed placeholders:', Array.from(new Set(placeholders)).slice(0, 15));
      }
    } else if (ext === '.pptx') {
      const zip = new AdmZip(fullPath);
      const entries = zip.getEntries();
      const slideEntries = entries.filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'));
      console.log(`Slides count: ${slideEntries.length}`);
      for (const slide of slideEntries) {
        const xml = zip.readAsText(slide);
        const matches = xml.match(/<a:t>([^<]+)<\/a:t>/g) || [];
        const text = matches.map(m => m.replace(/<[^>]+>/g, '')).join(' ');
        console.log(`  ${slide.entryName}: ${text.substring(0, 200)}`);
      }
    }
  }
}

inspectAll().catch(err => console.error(err));
