// Read-only live mirror of the Google Sheet database.
//
// - Downloads the sheet (all tabs, one request) and keeps the parsed result in memory.
// - Refreshes in the background every SYNC_INTERVAL_MS (default 60 s) and on demand (/api/refresh).
// - Requests never wait for Google: they always read the in-memory copy.
// - Nothing is ever written to the sheet or to disk.
// - If a refresh fails, the last good copy keeps being served.
const crypto = require('crypto');
const fs = require('fs');
const XLSX = require('xlsx');
const CONFIG = require('./config');

const SHEET_ID = process.env.DATABASE_SHEET_ID || '1gGptlk5dUXouGZT_mOlXu8xGzKmmrCGRSLjoDlO6ISM';
const SYNC_INTERVAL_MS = Number(process.env.SYNC_INTERVAL_MS) || 60 * 1000;
const MIN_GAP_MS = 2000;            // never hit Google more often than this
const FILE_PATH = process.env.DB_PATH || ''; // optional: read a local .xlsx instead of Google
const TABS = ['Cabinet_Data', 'Laminate_Data', 'masterdata'];

const state = {
  db: null,
  hash: '',
  lastCheck: 0,      // last time we asked Google
  lastChange: 0,     // last time the data actually changed
  lastError: '',
  inflight: null
};

function rows(wb, name) {
  const ws = wb.Sheets[name];
  return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true }) : null;
}

function parse(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', sheets: TABS, cellFormula: false, cellHTML: false, cellStyles: false });
  const db = { cabinets: {}, laminates: {}, accessories: [] };

  const cab = rows(wb, 'Cabinet_Data');
  if (cab) {
    for (let i = 1; i < cab.length; i++) {
      const name = String(cab[i][0] || '').toLowerCase().trim();
      const code = String(cab[i][1] || '').toUpperCase().trim();
      if (name && code) db.cabinets[name] = code;
    }
  }

  const lam = rows(wb, 'Laminate_Data');
  if (!lam) throw new Error("Cannot find a tab named 'Laminate_Data' in the database sheet.");
  for (let i = 1; i < lam.length; i++) {
    const key = String(lam[i][0] || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    if (!key) continue;
    db.laminates[key] = {
      code: String(lam[i][1] || '').trim(),
      url: String(lam[i][2] || '').trim(),
      edgeBand: String(lam[i][3] || '').trim(),
      type: String(lam[i][4] || '').trim()
    };
  }

  const master = rows(wb, 'masterdata');
  if (master) {
    for (let i = 1; i < master.length; i++) {
      const acc = String(master[i][0] || '').toLowerCase().trim();
      if (acc) db.accessories.push(acc);
    }
  }
  if (db.accessories.length === 0) db.accessories = CONFIG.ACCESSORY_ITEMS;
  return db;
}

async function download() {
  if (FILE_PATH) return fs.readFileSync(FILE_PATH);
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;
  const res = await fetch(url, { redirect: 'follow', signal: AbortSignal.timeout(20000) });
  const type = res.headers.get('content-type') || '';
  if (!res.ok || type.includes('text/html')) {
    throw new Error(
      `Google refused the download (HTTP ${res.status}). Share the database sheet as ` +
      `"Anyone with the link: Viewer" so the server can read it.`
    );
  }
  return Buffer.from(await res.arrayBuffer());
}

// Single-flight refresh: concurrent callers share one download.
function refresh({ force = false } = {}) {
  if (state.inflight) return state.inflight;
  if (!force && Date.now() - state.lastCheck < MIN_GAP_MS) return Promise.resolve(status());

  state.inflight = (async () => {
    try {
      const buf = await download();
      state.lastCheck = Date.now();
      const hash = crypto.createHash('md5').update(buf).digest('hex');
      if (hash !== state.hash || !state.db) {
        const db = parse(buf);            // parse first; only swap on success
        state.db = db;
        state.hash = hash;
        state.lastChange = Date.now();
        console.log(`[db] synced: ${Object.keys(db.laminates).length} laminates, ${Object.keys(db.cabinets).length} cabinets, ${db.accessories.length} accessories`);
      }
      state.lastError = '';
    } catch (e) {
      state.lastCheck = Date.now();
      state.lastError = e.message;
      console.warn('[db] sync failed: ' + e.message + (state.db ? ' (serving last good copy)' : ''));
    } finally {
      state.inflight = null;
    }
    return status();
  })();
  return state.inflight;
}

function start() {
  const first = refresh({ force: true });
  setInterval(() => refresh({ force: true }), SYNC_INTERVAL_MS).unref();
  return first;
}

// Always instant: reads memory only.
function getDatabase() {
  if (!state.db) {
    throw new Error(state.lastError || 'Database is still loading, try again in a few seconds.');
  }
  return state.db;
}

function status() {
  return {
    ok: !!state.db,
    laminates: state.db ? Object.keys(state.db.laminates).length : 0,
    cabinets: state.db ? Object.keys(state.db.cabinets).length : 0,
    lastCheck: state.lastCheck,
    lastChange: state.lastChange,
    syncIntervalSec: Math.round(SYNC_INTERVAL_MS / 1000),
    error: state.lastError
  };
}

module.exports = { start, refresh, getDatabase, status };
