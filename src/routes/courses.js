// src/routes/courses.js - Course API routes
const express = require('express');
const db = require('../db');
const router = express.Router();

// GET /api/courses — filter by ?college_id= or ?q=
router.get('/', (req, res) => {
  try {
    const { college_id, q } = req.query;
    let sql = `
      SELECT courses.*, colleges.name AS college_name
      FROM courses JOIN colleges ON courses.college_id = colleges.id
      WHERE 1=1
    `;
    const params = [];
    if (college_id) { sql += ' AND courses.college_id = ?'; params.push(college_id); }
    if (q)          { sql += ' AND courses.course_name LIKE ?'; params.push(`%${q}%`); }
    sql += ' ORDER BY courses.course_name ASC';

    const rows = params.length
      ? db.prepare(sql).all(params)
      : db.prepare(sql).all([]);
    res.json(rows);
  } catch (err) {
    console.error('Error in GET /api/courses:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/courses — add course to a college
router.post('/', (req, res) => {
  try {
    const { college_id, course_name } = req.body;
    if (!college_id || !course_name)
      return res.status(400).json({ error: 'college_id and course_name are required' });
    const result = db.prepare(
      'INSERT INTO courses (college_id, course_name) VALUES (?, ?)'
    ).run([college_id, course_name]);
    res.status(201).json({ id: result.lastInsertRowid, college_id, course_name });
  } catch (err) {
    console.error('Error in POST /api/courses:', err);
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;