(function () {
  const $ = id => document.getElementById(id);
  const drop = $('drop'), fileInput = $('file'), statusEl = $('status'), resultEl = $('result'), dbInfo = $('dbinfo');
  const printBtn = $('printBtn'), againBtn = $('againBtn'), imgPick = $('imgPick'), shell = $('shell');

  let lastMatrix = null;   // the uploaded estimate stays in memory only
  let lastResult = null;
  let pending = null;      // which manual image box is waiting for a file
  const manual = Object.create(null);   // room name -> manually added rows (survive sheet re-syncs)
  const editing = new Set();            // room names currently in "add item" mode
  const origTitle = document.title;

  // Glass cabinets: the shutter finish is left empty; the user types text and/or picks a glass type.
  const GLASS = [
    { id: 'fluted', name: 'Fluted Glass', src: 'glass/fluted.jpg' },
    { id: 'black', name: 'Black Tinted Glass', src: 'glass/black.jpg' },
    { id: 'grey', name: 'Grey Tinted Glass', src: 'glass/grey.jpg' },
    { id: 'clear', name: 'Clear Glass', src: 'glass/clear.jpg' },
    { id: 'brown', name: 'Brown Tinted Glass', src: 'glass/brown.jpg' }
  ];
  const glassById = Object.fromEntries(GLASS.map(g => [g.id, g]));
  const glassState = Object.create(null);   // "room||product||n" -> { type, text } (survives sheet re-syncs)
  function glassKeys(room) {
    const seen = {};
    return room.rows.map(r => {
      if (!r.glass) return null;
      seen[r.product] = (seen[r.product] || 0) + 1;
      return room.name + '||' + r.product + '||' + seen[r.product];
    });
  }
  const autosize = t => { t.style.height = 'auto'; t.style.height = (t.scrollHeight + 2) + 'px'; };

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
      const um = lastResult.unmatched.length;
      setStatus((reason === 'sheet' ? `Sheet changed: results updated automatically (${ms} ms)` : `Done in ${ms} ms`)
        + (um ? ` · ⚠ ${um} finish${um === 1 ? '' : 'es'} not found in the database` : ''), !!um);
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
      for (const k in glassState) delete glassState[k];
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
      name = `<div class="fin-name${isFill ? ' fill' : ''}${f.missing ? ' miss' : ''}">${esc(f.name)}${f.type ? ` <b>(${esc(f.type)})</b>` : ''}</div>`;
      if (f.missing) name += '<div class="miss-note">⚠ Not found in database</div>';
    }
    const code = f.code ? `<div class="fin-code">Code: ${esc(f.code)}</div>` : '';
    const edge = f.edge ? `<div class="fin-edge">Edge Band: ${esc(f.edge)}</div>` : '';
    return `<div class="fin"><div class="fin-text">${name}${code}${edge}</div>${swatch(f.image, !!f.name)}</div>`;
  }

  function glassCell(key) {
    const g = glassState[key] || {};
    const t = glassById[g.type];
    const opts = GLASS.map(o => `<option value="${o.id}"${o.id === g.type ? ' selected' : ''}>${esc(o.name)}</option>`).join('');
    return `<div class="fin">
      <div class="fin-text"><textarea class="gtxt" rows="2" data-gk="${esc(key)}" placeholder="Shutter finish (type, or pick a glass)">${esc(g.text || '')}</textarea></div>
      <div class="gbox${t ? ' has' : ''}" title="Choose glass type">${t ? `<img src="${t.src}" alt="${esc(t.name)}">` : '<span>Select<br>glass ▾</span>'}
        <select class="gsel" data-gk="${esc(key)}" aria-label="Glass type"><option value="">Select glass…</option>${opts}</select>
      </div></div>`;
  }

  const extractedRow = (r, gk) => `<tr><td class="prod">${esc(r.product)}</td><td>${esc(r.carcassMaterial)}</td><td>${esc(r.shutterMaterial)}</td>
    <td>${finishCell(r.carcass)}</td><td>${r.glass ? glassCell(gk) : finishCell(r.shutter)}</td></tr>`;

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
    shell.classList.toggle('has-results', !!lastMatrix && !!lastResult);
    printBtn.hidden = !has;
    againBtn.hidden = !lastMatrix;
    if (!lastResult) { resultEl.innerHTML = ''; return; }
    if (!has) { resultEl.innerHTML = '<p style="color:var(--muted);margin-top:24px">No materials found in this estimate.</p>'; return; }

    const um = lastResult.unmatched || [];
    const banner = um.length ? `<div class="warn no-print"><b>⚠ ${um.length} finish${um.length === 1 ? '' : 'es'} in this estimate ${um.length === 1 ? 'is' : 'are'} not in the database</b>
      <span>They are highlighted in red below. Add ${um.length === 1 ? 'it' : 'them'} to the Laminate_Data sheet and this page updates by itself within a minute.</span>
      <ul>${um.map(u => `<li><b>${esc(u.name)}</b> <small>${esc(u.rooms.join(', '))}</small></li>`).join('')}</ul></div>` : '';
    resultEl.innerHTML = banner + lastResult.rooms.map((room, ri) => {
      const isEd = editing.has(room.name);
      const rows = (manual[room.name] || []).map((r, i) => isEd ? manualEdit(r, i) : manualStatic(r)).join('');
      const gkeys = glassKeys(room);
      return `<section class="room" data-room="${ri}">
        <div class="room-head">${esc(room.name)}</div>
        <table>
          <colgroup><col class="c-prod"><col class="c-mat"><col class="c-mat"><col class="c-fin"><col class="c-fin"></colgroup>
          <thead><tr><th>Product</th><th>Carcass Mat</th><th>Shutter Mat</th><th>Carcass Finish</th><th>Shutter Finish</th></tr></thead>
          <tbody>${room.rows.map((r, i) => extractedRow(r, gkeys[i])).join('')}${rows}</tbody>
        </table>
        <div class="actions no-print"><button class="btn-pill" data-act="add">+ Add Manual Item</button>${isEd ? '<button class="btn-done" data-act="done">Done</button>' : ''}</div>
      </section>`;
    }).join('');
    resultEl.querySelectorAll('.gtxt').forEach(autosize);
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

  // glass: typing edits the text; choosing a glass type swaps the image and fills in its name
  resultEl.addEventListener('input', e => {
    const t = e.target;
    if (!t.classList || !t.classList.contains('gtxt')) return;
    const g = glassState[t.dataset.gk] || (glassState[t.dataset.gk] = { type: '', text: '' });
    g.text = t.value;
    autosize(t);
  });
  resultEl.addEventListener('change', e => {
    const sel = e.target;
    if (!sel.classList || !sel.classList.contains('gsel')) return;
    const g = glassState[sel.dataset.gk] || (glassState[sel.dataset.gk] = { type: '', text: '' });
    const t = glassById[sel.value];
    g.type = t ? t.id : '';
    if (t) g.text = t.name;
    render();
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

  // ---------- PDF: built in the page, always A4 landscape, downloaded directly ----------
  const fin = f => ({ name: f.name, type: f.type, code: f.code, edge: f.edge, img: f.image, box: !!f.name, missing: !!f.missing });
  function pdfRooms() {
    return lastResult.rooms.map(room => {
      const gkeys = glassKeys(room);
      const rows = room.rows.map((r, i) => {
        let s = fin(r.shutter);
        if (r.glass) {
          const g = glassState[gkeys[i]] || {};
          const t = glassById[g.type];
          s = { text: g.text || '', img: t ? t.src : '', box: true };
        }
        return { product: r.product, cMat: r.carcassMaterial, sMat: r.shutterMaterial, cFin: fin(r.carcass), sFin: s };
      });
      (manual[room.name] || []).forEach(m => {
        if (isEmptyRow(m)) return;
        rows.push({ product: m.product, cMat: m.cMat, sMat: m.sMat,
          cFin: { text: m.cFin, img: m.cImg, box: !!(m.cFin || m.cImg) },
          sFin: { text: m.sFin, img: m.sImg, box: !!(m.sFin || m.sImg) } });
      });
      return { name: room.name, rows };
    });
  }

  printBtn.addEventListener('click', async () => {
    if (!lastResult) return;
    finishAll();
    render();
    const title = (lastResult.projectName || 'Materials').replace(/[\\/:*?"<>|]+/g, ' ').trim() || 'Materials';
    printBtn.disabled = true;
    setStatus('Preparing PDF…');
    try {
      const out = await FB_pdf.create(pdfRooms(), { title });
      FB_pdf.download(out.blob, title + '.pdf');
      setStatus(out.missing
        ? `PDF downloaded (A4 landscape), but ${out.missing} swatch image(s) could not be added because the image host blocked them. They show as empty dotted boxes.`
        : `PDF downloaded: ${title}.pdf (A4 landscape, ${out.pages} page${out.pages === 1 ? '' : 's'})`, !!out.missing);
    } catch (e) {
      console.error(e);
      setStatus('Could not build the PDF directly (' + e.message + '). Opening the browser print dialog instead: choose A4, Landscape, and turn off Headers and footers.', true);
      document.title = title;
      setTimeout(() => window.print(), 50);
    } finally {
      printBtn.disabled = false;
    }
  });
  window.addEventListener('afterprint', () => { document.title = origTitle; });

  // ---------- wiring ----------
  drop.addEventListener('click', () => fileInput.click());
  againBtn.addEventListener('click', () => fileInput.click());
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
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
