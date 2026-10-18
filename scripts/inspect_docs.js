// scripts/inspect_docs.js
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

const dir = 'C:\\Users\\shlok\\Downloads\\Web_Development__Deliverables\\Deliverables';

const docxFiles = fs.readdirSync(dir).filter(f => f.endsWith('.docx') || f.endsWith('.pptx'));
for (const file of docxFiles) {
  const full = path.join(dir, file);
  const zip = new AdmZip(full);
  console.log(`\n========================================`);
  console.log(`FILE: ${file}`);
  console.log(`========================================`);
  const entries = zip.getEntries();
  const xmlEntries = entries.filter(e => e.entryName === 'word/document.xml' || e.entryName.startsWith('ppt/slides/slide'));
  for (const entry of xmlEntries) {
    const xml = zip.readAsText(entry);
    // Find text content inside <w:t> or <a:t>
    const matches = xml.match(/<(?:w|a):t[^>]*>([^<]+)<\/(?:w|a):t>/g) || [];
    const plain = matches.map(m => m.replace(/<[^>]+>/g, '')).join(' ');
    console.log(`Entry: ${entry.entryName}`);
    console.log(`Text snippet: ${plain.substring(0, 400)}\n`);
  }
}
