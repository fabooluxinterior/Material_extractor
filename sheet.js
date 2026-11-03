// Read-only live mirror of the Google Sheet database, running in the browser.
// - Reads Cabinet_Data, Laminate_Data and masterdata with the Google Sheets API (read-only key).
// - Keeps the parsed result in memory and in this browser's localStorage (instant start next visit).
// - Re-checks every SYNC_MS, when the tab becomes visible again, and on demand.
// - Never writes anything to the sheet.
(function () {
  const S = globalThis.FB_SETTINGS || {};
  const CFG = globalThis.FB_CONFIG || (typeof require === 'function' ? require('./config') : {});
  const TABS = ['Cabinet_Data', 'Laminate_Data', 'masterdata'];
  const LS_KEY = 'fb_sheet_cache_v1';
  const SYNC_MS = S.SYNC_MS || 60000;
  const MIN_GAP = 2000;

  const state = { db: null, sig: '', lastCheck: 0, lastChange: 0, error: '', cached: false };
  const subs = [];
  let inflight = null;
  let timer = null;

  const configured = () => !!(S.SHEET_ID && S.API_KEY);

  function emit(changed) {
    const s = status();
    subs.forEach(f => { try { f(s, changed); } catch (e) { /* ignore UI errors */ } });
  }

  function lsGet() { try { return JSON.parse(localStorage.getItem(LS_KEY) || 'null'); } catch (e) { return null; } }
  function lsSet(v) { try { localStorage.setItem(LS_KEY, JSON.stringify(v)); } catch (e) { /* storage full/blocked: fine */ } }

  // Same parsing rules as the original Apps Script fetchDatabaseMappings().
  function buildDb(raw) {
    const db = { cabinets: {}, laminates: {}, accessories: [] };

    const cab = raw.Cabinet_Data;
    if (cab) {
      for (let i = 1; i < cab.length; i++) {
        const r = cab[i] || [];
        const name = String(r[0] || '').toLowerCase().trim();
        const code = String(r[1] || '').toUpperCase().trim();
        if (name && code) db.cabinets[name] = code;
      }
    }

    const lam = raw.Laminate_Data;
    if (!lam) throw new Error("Cannot find a tab named 'Laminate_Data' in your Database sheet.");
    for (let i = 1; i < lam.length; i++) {
      const r = lam[i] || [];
      const key = String(r[0] || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
      if (!key) continue;
      db.laminates[key] = {
        code: String(r[1] || '').trim(),
        url: String(r[2] || '').trim(),
        edgeBand: String(r[3] || '').trim(),
        type: String(r[4] || '').trim()
      };
    }

    const master = raw.masterdata;
    if (master) {
      for (let i = 1; i < master.length; i++) {
        const acc = String((master[i] || [])[0] || '').toLowerCase().trim();
        if (acc) db.accessories.push(acc);
      }
    }
    if (db.accessories.length === 0) db.accessories = CFG.ACCESSORY_ITEMS;
    return db;
  }

  const BASE = () => `https://sheets.googleapis.com/v4/spreadsheets/${S.SHEET_ID}/values`;

  async function call(url) {
    const res = await fetch(url, { cache: 'no-store' });
    if (res.ok) return res.json();
    let detail = '';
    try { detail = (await res.json()).error.message || ''; } catch (e) { /* no body */ }
    const err = new Error(detail || ('HTTP ' + res.status));
    err.status = res.status;
    err.detail = detail;
    throw err;
  }

  const isMissingTab = e => e.status === 400 && !/api key/i.test(e.detail || '');

  async function fetchRaw() {
    const raw = {};
    try {
      // One request for all three tabs.
      const url = `${BASE()}:batchGet?${TABS.map(t => 'ranges=' + encodeURIComponent(t)).join('&')}&key=${encodeURIComponent(S.API_KEY)}`;
      const j = await call(url);
      TABS.forEach((t, i) => { raw[t] = (j.valueRanges[i] && j.valueRanges[i].values) || []; });
    } catch (e) {
      if (!isMissingTab(e)) throw e;
      // A tab is missing: read them one by one (Cabinet_Data and masterdata are optional).
      for (const t of TABS) {
        try {
          const j = await call(`${BASE()}/${encodeURIComponent(t)}?key=${encodeURIComponent(S.API_KEY)}`);
          raw[t] = j.values || [];
        } catch (e2) {
          if (isMissingTab(e2)) raw[t] = null; else throw e2;
        }
      }
    }
    return raw;
  }

  function explain(e) {
    if (e.status === 403) return 'Google refused access (403). Check: sheet shared as "Anyone with the link: Viewer", Google Sheets API enabled, and the API key allows this website.';
    if (e.status === 400 && /api key/i.test(e.detail || '')) return 'The API key in settings.js is not valid.';
    if (e.status === 404) return 'Database sheet not found. Check SHEET_ID in settings.js.';
    if (e instanceof TypeError) return 'Cannot reach Google (offline?). Using the last saved copy.';
    return e.message || String(e);
  }

  function refresh(opts) {
    const force = !!(opts && opts.force);
    if (inflight) return inflight;
    if (!force && Date.now() - state.lastCheck < MIN_GAP) return Promise.resolve(status());
    if (!configured()) {
      state.error = 'Setup needed: add your Google API key in settings.js (see README).';
      emit(false);
      return Promise.resolve(status());
    }

    inflight = (async () => {
      let changed = false;
      try {
        const raw = await fetchRaw();
        state.lastCheck = Date.now();
        const sig = JSON.stringify(raw);
        if (sig !== state.sig || !state.db) {
          const db = buildDb(raw);          // parse first; only swap on success
          state.db = db;
          state.sig = sig;
          state.lastChange = Date.now();
          state.cached = false;
          changed = true;
          lsSet({ sig, lastChange: state.lastChange });
        }
        state.error = '';
      } catch (e) {
        state.lastCheck = Date.now();
        state.error = explain(e);
      } finally {
        inflight = null;
      }
      emit(changed);
      return status();
    })();
    return inflight;
  }

  function start() {
    // Instant start from the last saved copy, then verify with Google.
    const saved = lsGet();
    if (saved && saved.sig) {
      try {
        state.db = buildDb(JSON.parse(saved.sig));
        state.sig = saved.sig;
        state.lastChange = saved.lastChange || 0;
        state.cached = true;
        emit(true);
      } catch (e) { /* ignore a bad cache */ }
    }
    refresh({ force: true });
    if (!timer) timer = setInterval(() => refresh({ force: true }), SYNC_MS);
    if (typeof document !== 'undefined') {
      document.addEventListener('visibilitychange', () => { if (!document.hidden) refresh({ force: true }); });
      window.addEventListener('online', () => refresh({ force: true }));
      window.addEventListener('focus', () => refresh());
    }
  }

  function getDatabase() {
    if (!state.db) throw new Error(state.error || 'Database is still loading, one moment…');
    return state.db;
  }

  function status() {
    return {
      ok: !!state.db,
      configured: configured(),
      laminates: state.db ? Object.keys(state.db.laminates).length : 0,
      cabinets: state.db ? Object.keys(state.db.cabinets).length : 0,
      lastCheck: state.lastCheck,
      lastChange: state.lastChange,
      cached: state.cached,
      error: state.error
    };
  }

  const exported = { start, refresh, getDatabase, status, buildDb, onChange: f => subs.push(f) };
  if (typeof module !== 'undefined' && module.exports) module.exports = exported;
  else globalThis.FB_sheet = exported;
})();
