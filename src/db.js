// src/db.js - Database setup using node-sqlite3-wasm (pure WASM, no compilation)
const { Database } = require('node-sqlite3-wasm');
const path = require('path');
const fs = require('fs');

const DB_PATH = path.join(__dirname, '..', 'data', 'collegekhoj.db');

// Ensure data directory exists
fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });

const db = new Database(DB_PATH);

// Enable WAL mode and foreign keys via exec (no .pragma() in wasm build)
db.exec('PRAGMA journal_mode=WAL');
db.exec('PRAGMA foreign_keys=ON');

// Create tables
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

module.exports = db;