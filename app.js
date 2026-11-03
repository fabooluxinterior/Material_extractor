(function () {
  const $ = id => document.getElementById(id);
  const drop = $('drop'), fileInput = $('file'), statusEl = $('status'), resultEl = $('result'), dbInfo = $('dbinfo');
  const printBtn = $('printBtn'), againBtn = $('againBtn'), imgPick = $('imgPick');

  let lastMatrix = null;   // the uploaded estimate stays in memory only
  let lastResult = null;
  let pending = null;      // which manual image box is waiting for a file
  const manual = Object.create(null);   // room name -> manually added rows (survive sheet re-syncs)
  const editing = new Set();            // room names currently in "add item" mode
  const origTitle = document.title;

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function setStatus(msg, err) { statusEl.textContent = msg || ''; statusEl.className = err ? 'err' : ''; }

  // ---------- sheet sync status ----------
  function paintSync() {
    const s = FB_sheet.status();
    if (!s.configured) { dbInfo.textContent = 'Setup needed: add API key in settings.js'; dbInfo.style.color = '#f87171'; return; }
    if (!s.ok) { dbInfo.textContent = s.error || 'Loading sheet…'; dbInfo.style.color = s.error ? '#f87171' : ''; return; }
    const ago = s.lastCheck ? Math.max(0, Math.round((Date.now() - s.lastCheck) / 1000)) + 's ago' : 'saved copy';
    dbInfo.textContent = `Sheet synced ${ago} · ${s.laminates} laminates` + (s.error ? ' · ' + s.error : '');
    dbInfo.style.color = s.error ? '#f87171' : '';
  }

  // ---------- extraction (instant, all local) ----------
  function extract(reason) {
    if (!lastMatrix) return;
    const s = FB_sheet.status();
    if (!s.ok) { setStatus(s.error || 'Loading sheet data…', !!s.error); return; }
    const t0 = performance.now();
    try {
      lastResult = FB_extractor.processMatrix(lastMatrix, FB_sheet.getDatabase());
      render();
      const ms = Math.round(performance.now() - t0);
      setStatus(reason === 'sheet' ? `Sheet changed: results updated automatically (${ms} ms)` : `Done in ${ms} ms`);
    } catch (e) {
      lastResult = null;
      render();
      setStatus(e.message, true);
    }
  }

  async function handleFile(file) {
    setStatus('Reading…');
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellFormula: false, cellHTML: false, cellStyles: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      lastMatrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
      for (const k in manual) delete manual[k];   // a new estimate starts clean
      editing.clear();
      extract('upload');
      FB_sheet.refresh();   // pick up any very recent sheet edit; re-renders by itself if something changed
    } catch (e) {
      setStatus('Could not read that file: ' + e.message, true);
    }
  }

  // ---------- rendering ----------
  function swatch(img, hasText) {
    if (img) return `<img class="sw" loading="lazy" referrerpolicy="no-referrer" src="${esc(img)}" alt="" onerror="this.replaceWith(Object.assign(document.createElement('div'),{className:'sw'}))">`;
    return hasText ? '<div class="sw"></div>' : '';
  }

  function finishCell(f) {
    let name = '';
    if (f.name) {
      const isFill = /fill manually/i.test(f.name);
      name = `<div class="fin-name${isFill ? ' fill' : ''}">${esc(f.name)}${f.type ? ` <b>(${esc(f.type)})</b>` : ''}</div>`;
    }
    const code = f.code ? `<div class="fin-code">Code: ${esc(f.code)}</div>` : '';
    const edge = f.edge ? `<div class="fin-edge">Edge Band: ${esc(f.edge)}</div>` : '';
    return `<div class="fin"><div class="fin-text">${name}${code}${edge}</div>${swatch(f.image, !!f.name)}</div>`;
  }

  const extractedRow = r => `<tr><td class="prod">${esc(r.product)}</td><td>${esc(r.carcassMaterial)}</td><td>${esc(r.shutterMaterial)}</td>
    <td>${finishCell(r.carcass)}</td><td>${finishCell(r.shutter)}</td></tr>`;

  function manualFin(text, img) {
    return `<div class="fin"><div class="fin-text"><div class="fin-name" style="white-space:pre-line">${esc(text)}</div></div>${swatch(img, !!text)}</div>`;
  }
  const manualStatic = r => `<tr><td class="prod">${esc(r.product)}</td><td>${esc(r.cMat)}</td><td>${esc(r.sMat)}</td>
    <td>${manualFin(r.cFin, r.cImg)}</td><td>${manualFin(r.sFin, r.sImg)}</td></tr>`;

  function imgBox(field, img) {
    return `<div class="imgbox" data-act="pick" data-f="${field}">${img ? `<img src="${img}" alt=""><span class="rm" data-act="imgrm" data-f="${field}">✕</span>` : '+ Image'}</div>`;
  }
  const manualEdit = (r, i) => `<tr data-i="${i}">
    <td><input class="inp b" data-f="product" value="${esc(r.product)}" placeholder="New Product"><button class="btn-del" data-act="del">Delete Row</button></td>
    <td><input class="inp" data-f="cMat" value="${esc(r.cMat)}" placeholder="Mat"></td>
    <td><input class="inp" data-f="sMat" value="${esc(r.sMat)}" placeholder="Mat"></td>
    <td><div class="fin"><input class="inp" data-f="cFin" value="${esc(r.cFin)}" placeholder="Finish Name -- Code">${imgBox('cImg', r.cImg)}</div></td>
    <td><div class="fin"><input class="inp" data-f="sFin" value="${esc(r.sFin)}" placeholder="Finish Name -- Code">${imgBox('sImg', r.sImg)}</div></td></tr>`;

  function render() {
    const has = !!(lastResult && lastResult.rooms.length);
    drop.hidden = !!lastMatrix && !!lastResult;
    printBtn.hidden = !has;
    againBtn.hidden = !lastMatrix;
    if (!lastResult) { resultEl.innerHTML = ''; return; }
    if (!has) { resultEl.innerHTML = '<p style="color:var(--muted);margin-top:24px">No materials found in this estimate.</p>'; return; }

    resultEl.innerHTML = lastResult.rooms.map((room, ri) => {
      const isEd = editing.has(room.name);
      const rows = (manual[room.name] || []).map((r, i) => isEd ? manualEdit(r, i) : manualStatic(r)).join('');
      return `<section class="room" data-room="${ri}">
        <div class="room-head">${esc(room.name)}</div>
        <table>
          <colgroup><col class="c-prod"><col class="c-mat"><col class="c-mat"><col class="c-fin"><col class="c-fin"></colgroup>
          <thead><tr><th>Product</th><th>Carcass Mat</th><th>Shutter Mat</th><th>Carcass Finish</th><th>Shutter Finish</th></tr></thead>
          <tbody>${room.rows.map(extractedRow).join('')}${rows}</tbody>
        </table>
        <div class="actions no-print"><button class="btn-pill" data-act="add">+ Add Manual Item</button>${isEd ? '<button class="btn-done" data-act="done">Done</button>' : ''}</div>
      </section>`;
    }).join('');
  }

  // ---------- manual items ----------
  const blank = () => ({ product: '', cMat: '', sMat: '', cFin: '', cImg: '', sFin: '', sImg: '' });
  const isEmptyRow = r => !(r.product || r.cMat || r.sMat || r.cFin || r.cImg || r.sFin || r.sImg);

  function finishEditing(name) {
    manual[name] = (manual[name] || []).filter(r => !isEmptyRow(r));
    editing.delete(name);
  }
  function finishAll() { [...editing].forEach(finishEditing); }

  function roomOf(el) {
    const sec = el.closest('.room');
    return sec && lastResult ? lastResult.rooms[+sec.dataset.room] : null;
  }

  // Shrink attached photos so they stay light (kept in memory only).
  function readImage(file) {
    return new Promise(resolve => {
      const fr = new FileReader();
      fr.onload = () => {
        const im = new Image();
        im.onload = () => {
          const max = 260, k = Math.min(1, max / Math.max(im.width, im.height));
          const c = document.createElement('canvas');
          c.width = Math.round(im.width * k); c.height = Math.round(im.height * k);
          c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
          resolve(c.toDataURL('image/jpeg', 0.88));
        };
        im.onerror = () => resolve('');
        im.src = fr.result;
      };
      fr.onerror = () => resolve('');
      fr.readAsDataURL(file);
    });
  }

  resultEl.addEventListener('click', e => {
    const el = e.target.closest('[data-act]');
    const room = el && roomOf(el);
    if (!room) return;
    const list = manual[room.name] || (manual[room.name] = []);
    const tr = el.closest('tr');
    const i = tr ? +tr.dataset.i : -1;

    switch (el.dataset.act) {
      case 'add':
        list.push(blank());
        editing.add(room.name);
        render();
        { const rows = resultEl.querySelectorAll(`.room[data-room="${[...resultEl.children].indexOf(el.closest('.room'))}"] tr[data-i] [data-f="product"]`);
          if (rows.length) rows[rows.length - 1].focus(); }
        break;
      case 'done': finishEditing(room.name); render(); break;
      case 'del': list.splice(i, 1); render(); break;
      case 'pick': pending = { room: room.name, i, f: el.dataset.f }; imgPick.click(); break;
      case 'imgrm': e.stopPropagation(); list[i][el.dataset.f] = ''; render(); break;
    }
  });

  resultEl.addEventListener('input', e => {
    const f = e.target.dataset && e.target.dataset.f;
    if (!f || e.target.type === 'file') return;
    const room = roomOf(e.target);
    const tr = e.target.closest('tr');
    if (room && tr && manual[room.name] && manual[room.name][+tr.dataset.i]) manual[room.name][+tr.dataset.i][f] = e.target.value;
  });

  imgPick.addEventListener('change', async () => {
    const file = imgPick.files[0];
    imgPick.value = '';
    if (!file || !pending) return;
    const data = await readImage(file);
    const row = manual[pending.room] && manual[pending.room][pending.i];
    if (row && data) { row[pending.f] = data; render(); }
    pending = null;
  });

  // ---------- print / PDF (A4 landscape, one room per page) ----------
  printBtn.addEventListener('click', () => {
    finishAll();
    render();
    document.title = (lastResult && lastResult.projectName) || 'Materials';
    setTimeout(() => window.print(), 50);   // let the re-render and images settle
  });
  window.addEventListener('afterprint', () => { document.title = origTitle; });

  // ---------- wiring ----------
  drop.addEventListener('click', () => fileInput.click());
  againBtn.addEventListener('click', () => fileInput.click());
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); if (!drop.hidden) drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  document.addEventListener('drop', e => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f && /\.(xlsx|xls|xlsm)$/i.test(f.name)) handleFile(f); });
  fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); fileInput.value = ''; });
  $('syncBtn').addEventListener('click', () => { dbInfo.textContent = 'Syncing…'; FB_sheet.refresh({ force: true }); });

  // Sheet changed (or first loaded): re-run the current estimate right away; manual items stay.
  FB_sheet.onChange((s, changed) => {
    paintSync();
    if (changed && lastMatrix) extract('sheet');
  });
  setInterval(paintSync, 1000);
  paintSync();
  render();
  FB_sheet.start();
})();
