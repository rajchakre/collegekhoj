// src/import.js - Import colleges from data/colleges.csv
const fs   = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');
const db   = require('./db');

const CSV_PATH = path.join(__dirname, '..', 'data', 'colleges.csv');

if (!fs.existsSync(CSV_PATH)) {
  console.error('ERROR: CSV file not found at', CSV_PATH);
  process.exit(1);
}

const csvContent = fs.readFileSync(CSV_PATH, 'utf8');
const records = parse(csvContent, {
  columns: true,
  skip_empty_lines: true,
  trim: true,
});

const insertCollege = db.prepare(
  'INSERT INTO colleges (name, city, state, contact_number) VALUES (?, ?, ?, ?)'
);
const insertCourse = db.prepare(
  'INSERT INTO courses (college_id, course_name) VALUES (?, ?)'
);

let collegeCount = 0;
let courseCount  = 0;

db.exec('BEGIN');
try {
  for (const row of records) {
    // node-sqlite3-wasm requires an ARRAY of params, not spread args
    const result = insertCollege.run([
      (row.name           || '').trim(),
      (row.city           || '').trim(),
      (row.state          || '').trim(),
      (row.contact_number || '').trim(),
    ]);

    const collegeId = result.lastInsertRowid;
    collegeCount++;

    const coursesRaw = (row.courses || '').trim();
    if (coursesRaw) {
      for (const course of coursesRaw.split(',').map(c => c.trim()).filter(Boolean)) {
        insertCourse.run([collegeId, course]);
        courseCount++;
      }
    }
  }
  db.exec('COMMIT');
  console.log(`Import complete: ${collegeCount} colleges, ${courseCount} courses.`);
} catch (err) {
  db.exec('ROLLBACK');
  console.error('Import failed:', err.message);
  process.exit(1);
}