const { DatabaseSync } = require('node:sqlite');
const path = require('path');

const db = new DatabaseSync(path.join(__dirname, 'withers.db'));

db.exec(`
  CREATE TABLE IF NOT EXISTS requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_name TEXT NOT NULL,
    phone TEXT NOT NULL,
    address TEXT NOT NULL,
    date TEXT NOT NULL,
    time TEXT NOT NULL,
    unit_count INTEGER NOT NULL,
    memo TEXT,
    status TEXT DEFAULT '접수',
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
  );
`);

module.exports = db;
