const path = require('path');
const express = require('express');
const compression = require('compression');
const multer = require('multer');
const XLSX = require('xlsx');
const db = require('./db');
const { processMatrix } = require('./extractor');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } }); // memory only, nothing saved
const REFRESH_TOKEN = process.env.REFRESH_TOKEN || '';
const STALE_MS = 10 * 1000; // an upload older-than-this triggers a background re-check

app.use(compression());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/extract', upload.single('file'), (req, res) => {
  const t0 = process.hrtime.bigint();
  try {
    if (!req.file) throw new Error('No file uploaded.');
    const database = db.getDatabase();            // memory read, instant
    const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellFormula: false, cellHTML: false, cellStyles: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
    const result = processMatrix(matrix, database);
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    res.json({ status: 'success', message: 'Materials Extracted Successfully!', ms: Math.round(ms), ...result });
  } catch (e) {
    res.status(400).json({ status: 'error', message: e.message });
  }
  // Keep the next upload fresh without making this one wait.
  if (Date.now() - db.status().lastCheck > STALE_MS) db.refresh();
});

app.get('/api/health', (req, res) => res.json(db.status()));

// Instant sync trigger (UI button, or a Google Sheets onEdit hook). Read-only towards the sheet.
app.all('/api/refresh', async (req, res) => {
  if (REFRESH_TOKEN && req.query.token !== REFRESH_TOKEN) return res.status(403).json({ ok: false, message: 'Bad token.' });
  res.json(await db.refresh({ force: true }));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Faboolux Materials Extractor running on http://localhost:${PORT}`);
  db.start();
});
