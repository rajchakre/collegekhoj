// src/server.js - Express server entry point
const express = require('express');
const path = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, '..', 'public')));

app.use('/api/colleges', require('./routes/colleges'));
app.use('/api/courses',  require('./routes/courses'));

app.get('/api/stats', (req, res) => {
  try {
    const db = require('./db');
    const colleges = db.prepare('SELECT COUNT(*) AS count FROM colleges').get([]);
    const courses   = db.prepare('SELECT COUNT(*) AS count FROM courses').get([]);
    const states    = db.prepare('SELECT COUNT(DISTINCT state) AS count FROM colleges').get([]);
    res.json({
      colleges: colleges ? colleges.count : 0,
      courses: courses ? courses.count : 0,
      states: states ? states.count : 0
    });
  } catch (err) {
    console.error('Error in /api/stats:', err);
    res.status(500).json({ error: err.message, colleges: 0, courses: 0, states: 0 });
  }
});

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

// Centralized error handler
app.use((err, req, res, next) => {
  console.error('Unhandled server error:', err);
  res.status(500).json({ error: 'Internal server error', details: err.message });
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`\nCollegeKhoj running at http://localhost:${PORT}`);
    console.log(`API: http://localhost:${PORT}/api/colleges\n`);
  });
}

module.exports = app;