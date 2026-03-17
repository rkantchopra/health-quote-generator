const Database = require('better-sqlite3');
const path = require('path');

// ─── Firebase Admin (keeps quotes safe across Render restarts) ───────────────
let fbDb = null;
try {
  if (process.env.FIREBASE_PROJECT_ID) {
    const admin = require('firebase-admin');
    admin.initializeApp({
      credential: admin.credential.cert({
        projectId: process.env.FIREBASE_PROJECT_ID,
        clientEmail: process.env.FIREBASE_CLIENT_EMAIL,
        privateKey: (process.env.FIREBASE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
      }),
      databaseURL: process.env.FIREBASE_DB_URL ||
        'https://incremint-dashboard-default-rtdb.asia-southeast1.firebasedatabase.app'
    });
    fbDb = admin.database();
    console.log('✅ Firebase persistence enabled');
  }
} catch (e) {
  console.warn('⚠️ Firebase init failed (quotes will be local-only):', e.message);
}

// ─── Local SQLite (fast, used for all read/write operations) ─────────────────
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

// ─── On startup: restore data from Firebase → SQLite ─────────────────────────
(async () => {
  if (!fbDb) return;
  try {
    // Restore quotes
    const qSnap = await fbDb.ref('health_quotes/quotes').once('value');
    const quotes = qSnap.val();
    if (quotes) {
      const stmt = db.prepare(`INSERT OR IGNORE INTO quotes
        (quote_id, customer_name, customer_phone, customer_email,
         advisor_id, advisor_name, submitted_by, members_json, insurers_json,
         advisor_note, reviewed, status, submitted_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
      let count = 0;
      for (const q of Object.values(quotes)) {
        try {
          stmt.run(q.quote_id, q.customer_name, q.customer_phone, q.customer_email,
            q.advisor_id, q.advisor_name, q.submitted_by, q.members_json, q.insurers_json,
            q.advisor_note, q.reviewed || 0, q.status || 'New', q.submitted_at);
          count++;
        } catch (e) { /* duplicate key — already exists, skip */ }
      }
      console.log(`✅ Restored ${count} quotes from Firebase`);
    }

    // Restore advisors
    const aSnap = await fbDb.ref('health_quotes/advisors').once('value');
    const advisors = aSnap.val();
    if (advisors) {
      const stmt = db.prepare(`INSERT OR IGNORE INTO advisors
        (id, name, email, password, created_at, is_active)
        VALUES (?, ?, ?, ?, ?, ?)`);
      let count = 0;
      for (const a of Object.values(advisors)) {
        try {
          stmt.run(a.id, a.name, a.email, a.password, a.created_at, a.is_active);
          count++;
        } catch (e) { /* skip */ }
      }
      console.log(`✅ Restored ${count} advisors from Firebase`);
    }
  } catch (e) {
    console.error('⚠️ Firebase restore failed:', e.message);
  }
})();

// ─── Firebase sync helpers (non-blocking, runs in background) ────────────────
function syncQuote(quoteId) {
  if (!fbDb || !quoteId) return;
  try {
    const q = db.prepare('SELECT * FROM quotes WHERE quote_id = ?').get(quoteId);
    if (q) {
      fbDb.ref(`health_quotes/quotes/${quoteId}`).set(q)
        .catch(e => console.error('FB quote sync error:', e.message));
    } else {
      fbDb.ref(`health_quotes/quotes/${quoteId}`).remove()
        .catch(e => console.error('FB quote remove error:', e.message));
    }
  } catch (e) { /* ignore */ }
}

function syncAdvisor(id) {
  if (!fbDb || !id) return;
  try {
    const a = db.prepare('SELECT * FROM advisors WHERE id = ?').get(id);
    if (a) {
      fbDb.ref(`health_quotes/advisors/${a.id}`).set(a)
        .catch(e => console.error('FB advisor sync error:', e.message));
    } else {
      fbDb.ref(`health_quotes/advisors/${id}`).remove()
        .catch(e => console.error('FB advisor remove error:', e.message));
    }
  } catch (e) { /* ignore */ }
}

// ─── Database API (same interface — server.js needs NO changes) ──────────────
const run = (sql, params = []) => {
  try {
    const stmt = db.prepare(sql);
    const result = stmt.run(...params);

    // Auto-sync mutations to Firebase (background, non-blocking)
    const s = sql.toLowerCase().trim();
    if (s.includes('quotes') && !s.startsWith('select')) {
      const quoteId = params.find(p => typeof p === 'string' && /^TK\d{8,}/.test(p));
      if (quoteId) syncQuote(quoteId);
    }
    if (s.includes('advisors') && !s.startsWith('select')) {
      const id = s.startsWith('insert') ? result.lastInsertRowid : params[params.length - 1];
      if (id) syncAdvisor(id);
    }

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
