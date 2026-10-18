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
  const db = require('./db');
  const colleges = db.prepare('SELECT COUNT(*) AS count FROM colleges').get([]);
  const courses   = db.prepare('SELECT COUNT(*) AS count FROM courses').get([]);
  const states    = db.prepare('SELECT COUNT(DISTINCT state) AS count FROM colleges').get([]);
  res.json({ colleges: colleges.count, courses: courses.count, states: states.count });
});

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`\nCollegeKhoj running at http://localhost:${PORT}`);
    console.log(`API: http://localhost:${PORT}/api/colleges\n`);
  });
}

module.exports = app;