// src/routes/colleges.js - College API routes
const express = require('express');
const db = require('../db');
const router = express.Router();

// GET /api/colleges  â€” list all, or search with ?q= &state= &city=
router.get('/', (req, res) => {
  const { q, state, city } = req.query;
  let sql    = 'SELECT * FROM colleges WHERE 1=1';
  const params = [];

  if (q) {
    sql += ' AND (name LIKE ? OR city LIKE ? OR state LIKE ?)';
    params.push(`%${q}%`, `%${q}%`, `%${q}%`);
  }
  if (state) { sql += ' AND state LIKE ?'; params.push(`%${state}%`); }
  if (city)  { sql += ' AND city LIKE ?';  params.push(`%${city}%`); }

  sql += ' ORDER BY name ASC';
  // node-sqlite3-wasm: pass array for params
  const colleges = params.length
    ? db.prepare(sql).all(params)
    : db.prepare(sql).all([]);
  res.json(colleges);
});

// GET /api/colleges/:id â€” single college + its courses
router.get('/:id', (req, res) => {
  const college = db.prepare('SELECT * FROM colleges WHERE id = ?').get([req.params.id]);
  if (!college) return res.status(404).json({ error: 'College not found' });
  const courses = db.prepare('SELECT * FROM courses WHERE college_id = ?').all([req.params.id]);
  res.json({ ...college, courses });
});

// POST /api/colleges â€” add new college
router.post('/', (req, res) => {
  const { name, city, state, contact_number } = req.body;
  if (!name) return res.status(400).json({ error: 'name is required' });
  const result = db.prepare(
    'INSERT INTO colleges (name, city, state, contact_number) VALUES (?, ?, ?, ?)'
  ).run([name, city || null, state || null, contact_number || null]);
  res.status(201).json({ id: result.lastInsertRowid, name, city, state, contact_number });
});

// DELETE /api/colleges/:id
router.delete('/:id', (req, res) => {
  const info = db.prepare('DELETE FROM colleges WHERE id = ?').run([req.params.id]);
  if (info.changes === 0) return res.status(404).json({ error: 'College not found' });
  res.json({ message: 'Deleted successfully' });
});

module.exports = router;