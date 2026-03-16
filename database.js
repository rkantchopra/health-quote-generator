const Database = require('better-sqlite3');
const path = require('path');

const db = new Database(path.join(__dirname, 'database.db'));

db.exec(`CREATE TABLE IF NOT EXISTS advisors (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT,
  email TEXT UNIQUE,
  password TEXT,
  created_at TEXT,
  is_active INTEGER DEFAULT 1
)`);

db.exec(`CREATE TABLE IF NOT EXISTS quotes (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  quote_id TEXT UNIQUE,
  customer_name TEXT,
  customer_phone TEXT,
  customer_email TEXT,
  advisor_id INTEGER,
  advisor_name TEXT,
  submitted_by TEXT,
  members_json TEXT,
  insurers_json TEXT,
  advisor_note TEXT,
  reviewed INTEGER DEFAULT 0,
  status TEXT DEFAULT 'New',
  submitted_at TEXT
)`);

const run = (sql, params = []) => {
  try {
    const stmt = db.prepare(sql);
    const result = stmt.run(...params);
    return Promise.resolve({ id: result.lastInsertRowid, changes: result.changes });
  } catch (err) {
    return Promise.reject(err);
  }
};

const get = (sql, params = []) => {
  try {
    const stmt = db.prepare(sql);
    return Promise.resolve(stmt.get(...params));
  } catch (err) {
    return Promise.reject(err);
  }
};

const all = (sql, params = []) => {
  try {
    const stmt = db.prepare(sql);
    return Promise.resolve(stmt.all(...params));
  } catch (err) {
    return Promise.reject(err);
  }
};

module.exports = { db, run, get, all };
