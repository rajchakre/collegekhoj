// src/db.js - Database setup using node-sqlite3-wasm (pure WASM, no compilation)
const { Database } = require('node-sqlite3-wasm');
const path = require('path');
const fs = require('fs');

let DB_PATH;

if (process.env.VERCEL) {
  // On Vercel serverless environment, only /tmp is writable
  const tmpDir = path.join('/tmp', 'collegekhoj-data');
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }
  const tmpDb = path.join(tmpDir, 'collegekhoj.db');
  const srcDb = path.join(__dirname, '..', 'data', 'collegekhoj.db');

  if (!fs.existsSync(tmpDb) && fs.existsSync(srcDb)) {
    try {
      fs.copyFileSync(srcDb, tmpDb);
    } catch (err) {
      console.warn('Could not copy sqlite db to /tmp:', err);
    }
  }
  DB_PATH = tmpDb;
} else {
  DB_PATH = path.join(__dirname, '..', 'data', 'collegekhoj.db');
  fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
}

// Clean up stale lock directory if left by a previous crash or abrupt exit
const lockPath = DB_PATH + '.lock';
if (fs.existsSync(lockPath)) {
  try {
    fs.rmSync(lockPath, { recursive: true, force: true });
  } catch (err) {}
}

const db = new Database(DB_PATH);

// Clean lock on process exit
const cleanLock = () => {
  try {
    if (fs.existsSync(lockPath)) {
      fs.rmSync(lockPath, { recursive: true, force: true });
    }
  } catch (e) {}
};
process.on('exit', cleanLock);
process.on('SIGINT', () => { cleanLock(); process.exit(0); });
process.on('SIGTERM', () => { cleanLock(); process.exit(0); });

// Configure busy timeout, journal mode, and foreign keys
try {
  db.exec('PRAGMA busy_timeout = 5000');
} catch (e) {}
try {
  db.exec('PRAGMA journal_mode = DELETE');
  db.exec('PRAGMA synchronous = NORMAL');
} catch (e) {}
try {
  db.exec('PRAGMA foreign_keys = ON');
} catch (e) {}

// Create tables if not already existing
try {
  db.exec(`
    CREATE TABLE IF NOT EXISTS colleges (
      id             INTEGER PRIMARY KEY AUTOINCREMENT,
      name           TEXT NOT NULL,
      city           TEXT,
      state          TEXT,
      contact_number TEXT
    );

    CREATE TABLE IF NOT EXISTS courses (
      id          INTEGER PRIMARY KEY AUTOINCREMENT,
      college_id  INTEGER NOT NULL,
      course_name TEXT NOT NULL,
      FOREIGN KEY (college_id) REFERENCES colleges(id) ON DELETE CASCADE
    );
  `);
} catch (err) {
  // When another process holds the DB, continue if already initialized
  console.warn('Database schema verification notice:', err.message);
}

module.exports = db;