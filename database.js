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
    unit_count INTEGER NOT NULL DEFAULT 1,
    ac_info TEXT DEFAULT '',
    memo TEXT,
    status TEXT DEFAULT '접수',
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
  );
`);

try {
  db.exec(`ALTER TABLE requests ADD COLUMN ac_info TEXT DEFAULT ''`);
} catch (e) {}

module.exports = db;
