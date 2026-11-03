(function () {
  const $ = id => document.getElementById(id);
  const drop = $('drop'), fileInput = $('file'), statusEl = $('status'), resultEl = $('result'), dbInfo = $('dbinfo');
  let lastMatrix = null;   // the uploaded estimate stays in memory only
  let lastResult = null;

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function setStatus(msg, err) { statusEl.textContent = msg; statusEl.className = err ? 'err' : ''; }

  // ---- sheet sync status ----
  function paintSync() {
    const s = FB_sheet.status();
    if (!s.configured) { dbInfo.textContent = 'Setup needed: add API key in settings.js'; dbInfo.style.color = 'var(--red)'; return; }
    if (!s.ok) { dbInfo.textContent = s.error || 'Loading sheet…'; dbInfo.style.color = s.error ? 'var(--red)' : 'var(--muted)'; return; }
    const ago = s.lastCheck ? Math.max(0, Math.round((Date.now() - s.lastCheck) / 1000)) + 's ago' : 'saved copy';
    dbInfo.textContent = `Sheet synced ${ago} · ${s.laminates} laminates` + (s.error ? ' · ' + s.error : '');
    dbInfo.style.color = s.error ? 'var(--red)' : 'var(--muted)';
  }

  // ---- extraction (instant, all local) ----
  function extract(reason) {
    if (!lastMatrix) return;
    const s = FB_sheet.status();
    if (!s.ok) { setStatus(s.error || 'Loading sheet data…', !!s.error); return; }
    const t0 = performance.now();
    try {
      lastResult = FB_extractor.processMatrix(lastMatrix, FB_sheet.getDatabase());
      render(lastResult);
      const ms = Math.round(performance.now() - t0);
      setStatus(reason === 'sheet' ? `Sheet changed: results updated automatically (${ms} ms)` : `Done in ${ms} ms`);
    } catch (e) {
      resultEl.innerHTML = '';
      setStatus(e.message, true);
    }
  }

  async function handleFile(file) {
    setStatus('Reading…');
    resultEl.innerHTML = '';
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellFormula: false, cellHTML: false, cellStyles: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      lastMatrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
      extract('upload');
      FB_sheet.refresh();   // pick up any very recent sheet edit; re-renders by itself if something changed
    } catch (e) {
      setStatus('Could not read that file: ' + e.message, true);
    }
  }

  // ---- rendering ----
  function finishCells(f) {
    const text = f.lines.join('\n');
    const isFill = /fill manually/i.test(text);
    const img = f.image ? `<img loading="lazy" referrerpolicy="no-referrer" src="${esc(f.image)}" alt="">` : '';
    return `<td class="finish${isFill ? ' fill' : ''}">${esc(text)}</td><td class="img">${img}</td>`;
  }

  function render(j) {
    const toolbar = `<div class="toolbar"><h2>${esc(j.projectName)}</h2>
      <button class="secondary" id="copyBtn">Copy as table</button></div>`;
    const rooms = j.rooms.map(room => `
      <div class="room"><h3>${esc(room.name)}</h3>
      <table>
        <tr><th>Product</th><th>Carcass Material</th><th>Shutter Material</th><th colspan="2">Carcass Finish</th><th colspan="2">Shutter Finish</th></tr>
        ${room.rows.map(r => `<tr>
          <td>${esc(r.product)}</td><td>${esc(r.carcassMaterial)}</td><td>${esc(r.shutterMaterial)}</td>
          ${finishCells(r.carcass)}${finishCells(r.shutter)}
        </tr>`).join('')}
      </table></div>`).join('');
    resultEl.innerHTML = toolbar + (rooms || '<p>No materials found.</p>');
    $('copyBtn').addEventListener('click', copyTable);
  }

  function copyTable() {
    if (!lastResult) return;
    const lines = [];
    lastResult.rooms.forEach(room => {
      lines.push(room.name);
      lines.push(['Product', 'Carcass Material', 'Shutter Material', 'Carcass Finish', 'Shutter Finish'].join('\t'));
      room.rows.forEach(r => lines.push([r.product, r.carcassMaterial, r.shutterMaterial,
        '"' + r.carcass.lines.join('\n') + '"', '"' + r.shutter.lines.join('\n') + '"'].join('\t')));
      lines.push('');
    });
    navigator.clipboard.writeText(lines.join('\n')).then(() => setStatus('Copied to clipboard'));
  }

  // ---- wiring ----
  drop.addEventListener('click', () => fileInput.click());
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => { if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); fileInput.value = ''; });
  $('syncBtn').addEventListener('click', () => { dbInfo.textContent = 'Syncing…'; FB_sheet.refresh({ force: true }); });

  // Sheet changed (or first loaded): re-run the current estimate right away.
  FB_sheet.onChange((s, changed) => {
    paintSync();
    if (changed && lastMatrix) extract('sheet');
  });
  setInterval(paintSync, 1000);
  paintSync();
  FB_sheet.start();
})();
