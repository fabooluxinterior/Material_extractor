// Builds the materials PDF entirely in the browser: real A4 LANDSCAPE pages, upright text,
// no print dialog and no browser header/footer. No libraries, nothing is uploaded anywhere.
(function () {
  const PW = 841.89, PH = 595.28;          // A4 landscape in points
  const M = 28;                            // page margin
  const PAD = 7;                           // cell padding
  const SW = 48;                           // swatch size
  const HDR = 24;                          // table header height
  const COLS = [0.15, 0.11, 0.11, 0.315, 0.315];
  const HEAD = ['PRODUCT', 'CARCASS MAT', 'SHUTTER MAT', 'CARCASS FINISH', 'SHUTTER FINISH'];
  const FONT = 'Arial, Helvetica, sans-serif';   // metric-compatible with the PDF's Helvetica

  const mctx = document.createElement('canvas').getContext('2d');
  function tw(s, size, bold) { mctx.font = (bold ? 'bold ' : '') + size + 'px ' + FONT; return mctx.measureText(s).width; }

  const n = v => String(Math.round(v * 100) / 100);
  const Y = y => PH - y;                    // we lay out from the top; PDF counts from the bottom
  const rgb = hex => { const v = parseInt(hex.slice(1), 16); return [(v >> 16) & 255, (v >> 8) & 255, v & 255].map(x => (x / 255).toFixed(3)).join(' '); };

  // ---- text encoding (WinAnsi) ----
  const WIN = { 0x2013: 0x96, 0x2014: 0x97, 0x2018: 0x91, 0x2019: 0x92, 0x201C: 0x93, 0x201D: 0x94, 0x2022: 0x95, 0x2026: 0x85, 0x20AC: 0x80, 0x2122: 0x99 };
  function latin(str) {
    let o = '';
    for (const ch of String(str == null ? '' : str)) {
      let c = ch.codePointAt(0);
      if (WIN[c]) c = WIN[c]; else if (c > 255) c = 63; else if (c < 32) c = 32;
      o += String.fromCharCode(c);
    }
    return o;
  }
  const esc = s => latin(s).replace(/[\\()]/g, m => '\\' + m);
  const s2b = str => { const u = new Uint8Array(str.length); for (let i = 0; i < str.length; i++) u[i] = str.charCodeAt(i) & 255; return u; };

  // ---- rounded-rectangle path (coordinates from the top-left) ----
  function rr(x, y, w, h, r) {
    const k = 0.5523 * r, X0 = x, X1 = x + w, Y0 = Y(y + h), Y1 = Y(y);
    return [
      `${n(X0 + r)} ${n(Y0)} m`, `${n(X1 - r)} ${n(Y0)} l`,
      `${n(X1 - r + k)} ${n(Y0)} ${n(X1)} ${n(Y0 + r - k)} ${n(X1)} ${n(Y0 + r)} c`, `${n(X1)} ${n(Y1 - r)} l`,
      `${n(X1)} ${n(Y1 - r + k)} ${n(X1 - r + k)} ${n(Y1)} ${n(X1 - r)} ${n(Y1)} c`, `${n(X0 + r)} ${n(Y1)} l`,
      `${n(X0 + r - k)} ${n(Y1)} ${n(X0)} ${n(Y1 - r + k)} ${n(X0)} ${n(Y1 - r)} c`, `${n(X0)} ${n(Y0 + r)} l`,
      `${n(X0)} ${n(Y0 + r - k)} ${n(X0 + r - k)} ${n(Y0)} ${n(X0 + r)} ${n(Y0)} c`, 'h'
    ].join('\n');
  }

  class Page {
    constructor() { this.ops = []; this.imgs = new Set(); }
    fillRect(x, y, w, h, hex) { this.ops.push(`q ${rgb(hex)} rg ${n(x)} ${n(Y(y + h))} ${n(w)} ${n(h)} re f Q`); }
    strokeRect(x, y, w, h, lw) { this.ops.push(`q 0 G ${n(lw)} w ${n(x)} ${n(Y(y + h))} ${n(w)} ${n(h)} re S Q`); }
    card(x, y, w, h, r, lw) { this.ops.push(`q 0 G ${n(lw)} w ${rr(x, y, w, h, r)} S Q`); }
    dotted(x, y, s, r) { this.ops.push(`q 0 G 1.1 w 1 J [0.01 2.8] 0 d ${rr(x, y, s, s, r)} S Q`); }
    image(k, x, y, s, r) { this.imgs.add(k); this.ops.push(`q ${rr(x, y, s, s, r)} W n ${n(s)} 0 0 ${n(s)} ${n(x)} ${n(Y(y + s))} cm /Im${k} Do Q`); }
    text(s, x, baseY, size, bold, gray) { this.ops.push(`BT /${bold ? 'F2' : 'F1'} ${n(size)} Tf ${n(gray || 0)} g ${n(x)} ${n(Y(baseY))} Td (${esc(s)}) Tj ET`); }
  }

  // ---- word wrapping with mixed regular/bold segments ----
  function wrap(segs, maxW, size) {
    const tokens = [];
    segs.forEach(s => {
      String(s.t == null ? '' : s.t).split('\n').forEach((part, pi) => {
        if (pi > 0) tokens.push({ nl: true });
        part.split(/\s+/).filter(Boolean).forEach(w => tokens.push({ t: w, b: !!s.b }));
      });
    });
    const lines = []; let cur = [], curW = 0;
    const sp = tw(' ', size, false);
    const flush = () => { lines.push(cur); cur = []; curW = 0; };
    tokens.forEach(tk => {
      if (tk.nl) { flush(); return; }
      let pieces = [tk.t];
      if (tw(tk.t, size, tk.b) > maxW) {          // very long word: break it
        pieces = []; let part = '';
        for (const ch of tk.t) { if (part && tw(part + ch, size, tk.b) > maxW) { pieces.push(part); part = ch; } else part += ch; }
        if (part) pieces.push(part);
      }
      pieces.forEach((p, i) => {
        const w = tw(p, size, tk.b);
        if (cur.length && curW + sp + w > maxW) flush();
        curW += (cur.length ? sp : 0) + w;
        cur.push({ t: p, b: tk.b });
        if (i < pieces.length - 1) flush();
      });
    });
    if (cur.length || !lines.length) flush();
    return lines;
  }

  function drawLines(pg, lines, x, top, size, lh, gray) {
    const sp = tw(' ', size, false);
    lines.forEach((ln, i) => {
      let cx = x; const base = top + i * lh + size * 0.84;
      ln.forEach((tk, j) => { if (j) cx += sp; pg.text(tk.t, cx, base, size, tk.b, gray); cx += tw(tk.t, size, tk.b); });
    });
  }

  // ---- images: load, square-crop (like the on-screen swatch), re-encode as JPEG ----
  function b64ToBytes(b64) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }
  function loadImage(src) {
    return new Promise(resolve => {
      const im = new Image();
      const timer = setTimeout(() => resolve(null), 10000);
      if (!/^(data|blob):/.test(src)) im.crossOrigin = 'anonymous';
      im.onload = () => {
        clearTimeout(timer);
        try {
          const side = Math.min(im.naturalWidth, im.naturalHeight), px = Math.min(240, side);
          const c = document.createElement('canvas'); c.width = c.height = px;
          const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, px, px);
          g.drawImage(im, (im.naturalWidth - side) / 2, (im.naturalHeight - side) / 2, side, side, 0, 0, px, px);
          resolve({ bytes: b64ToBytes(c.toDataURL('image/jpeg', 0.9).split(',')[1]), w: px, h: px });
        } catch (e) { resolve(null); }       // blocked by the image host (no CORS)
      };
      im.onerror = () => { clearTimeout(timer); resolve(null); };
      im.src = src;
    });
  }

  // ---- file assembly ----
  function assemble(pages, list, title) {
    const parts = []; let pos = 0; const offs = [];
    const put = x => { const u = typeof x === 'string' ? s2b(x) : x; parts.push(u); pos += u.length; };
    const obj = (num, str) => { offs[num] = pos; put(num + ' 0 obj\n' + str + '\nendobj\n'); };
    const stream = (num, dict, data) => { offs[num] = pos; put(num + ' 0 obj\n<< ' + dict + ' /Length ' + data.length + ' >>\nstream\n'); put(data); put('\nendstream\nendobj\n'); };

    let next = 6 + list.length;
    const nums = pages.map(() => { const c = next++, p = next++; return { c, p }; });
    const total = next;

    put('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');
    obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
    obj(2, `<< /Type /Pages /Count ${pages.length} /Kids [${nums.map(o => o.p + ' 0 R').join(' ')}] >>`);
    obj(3, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>');
    obj(4, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>');
    obj(5, `<< /Title (${esc(title)}) /Producer (Faboolux Materials Extractor) >>`);
    list.forEach((im, i) => stream(6 + i, `/Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode`, im.bytes));
    pages.forEach((pg, i) => {
      stream(nums[i].c, '', s2b(pg.ops.join('\n')));
      const xo = [...pg.imgs].map(k => `/Im${k} ${6 + k} 0 R`).join(' ');
      obj(nums[i].p, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PW} ${PH}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> /XObject << ${xo} >> >> /Contents ${nums[i].c} 0 R >>`);
    });
    const xref = pos;
    let x = 'xref\n0 ' + total + '\n0000000000 65535 f \n';
    for (let i = 1; i < total; i++) x += String(offs[i]).padStart(10, '0') + ' 00000 n \n';
    put(x + 'trailer\n<< /Size ' + total + ' /Root 1 0 R /Info 5 0 R >>\nstartxref\n' + xref + '\n%%EOF\n');
    return new Blob(parts, { type: 'application/pdf' });
  }

  // rooms: [{ name, rows: [{ product, cMat, sMat, cFin, sFin }] }]
  // a finish is { name, type, code, edge } or { text }, plus { img, box }
  async function create(rooms, opts) {
    const title = (opts && opts.title) || 'Materials';

    const srcs = [...new Set(rooms.flatMap(r => r.rows.flatMap(x => [x.cFin.img, x.sFin.img])).filter(Boolean))];
    const loaded = await Promise.all(srcs.map(loadImage));
    const idx = new Map(), list = []; let missing = 0;
    srcs.forEach((s, i) => { if (loaded[i]) { idx.set(s, list.length); list.push(loaded[i]); } else missing++; });

    const tx = M + 3, tWidth = PW - 2 * M - 6;
    const cw = COLS.map(c => c * tWidth);
    const colX = []; cw.reduce((acc, w, i) => { colX[i] = acc; return acc + w; }, tx);
    const LIMIT = PH - M - 10;

    function finishLayout(f, width) {
      const hasBox = !!(f.img || f.box);
      const textW = width - 2 * PAD - (hasBox ? SW + 8 : 0);
      const blocks = [];
      if (f.text) {
        blocks.push({ lines: wrap([{ t: f.text }], textW, 9), size: 9, lh: 11.8, gap: 0, gray: 0 });
      } else if (f.name) {
        const segs = [{ t: f.name }]; if (f.type) segs.push({ t: '(' + f.type + ')', b: true });
        blocks.push({ lines: wrap(segs, textW, 9), size: 9, lh: 11.8, gap: 0, gray: 0 });
        if (f.code) blocks.push({ lines: wrap([{ t: 'Code: ' + f.code, b: true }], textW, 7.6), size: 7.6, lh: 10, gap: 4, gray: 0 });
        if (f.edge) blocks.push({ lines: wrap([{ t: 'Edge Band: ' + f.edge }], textW, 7.6), size: 7.6, lh: 10, gap: 4, gray: 0.25 });
      }
      const th = blocks.reduce((a, b) => a + b.gap + b.lines.length * b.lh, 0);
      return { hasBox, blocks, f, h: Math.max(th, hasBox ? SW : 0) + 2 * PAD };
    }

    function measure(r) {
      const p = wrap([{ t: r.product, b: true }], cw[0] - 2 * PAD, 9);
      const c = wrap([{ t: r.cMat }], cw[1] - 2 * PAD, 9);
      const s = wrap([{ t: r.sMat }], cw[2] - 2 * PAD, 9);
      const f1 = finishLayout(r.cFin, cw[3]), f2 = finishLayout(r.sFin, cw[4]);
      return { p, c, s, f1, f2, h: Math.max(30, 2 * PAD + 11.8 * Math.max(p.length, c.length, s.length), f1.h, f2.h) };
    }

    function drawFinish(pg, L, x, y, width) {
      let ty = y + PAD;
      L.blocks.forEach(b => { ty += b.gap; drawLines(pg, b.lines, x + PAD, ty, b.size, b.lh, b.gray); ty += b.lines.length * b.lh; });
      if (L.hasBox) {
        const bx = x + width - PAD - SW, by = y + PAD;
        if (L.f.img && idx.has(L.f.img)) pg.image(idx.get(L.f.img), bx, by, SW, 6);
        pg.dotted(bx, by, SW, 6);
      }
    }

    function drawRow(pg, m, y) {
      for (let i = 0; i < 5; i++) pg.strokeRect(colX[i], y, cw[i], m.h, 0.7);
      drawLines(pg, m.p, colX[0] + PAD, y + PAD, 9, 11.8, 0);
      drawLines(pg, m.c, colX[1] + PAD, y + PAD, 9, 11.8, 0);
      drawLines(pg, m.s, colX[2] + PAD, y + PAD, 9, 11.8, 0);
      drawFinish(pg, m.f1, colX[3], y, cw[3]);
      drawFinish(pg, m.f2, colX[4], y, cw[4]);
    }

    const pages = [];
    rooms.forEach(room => {
      let pg, y;
      const start = () => {
        pg = new Page(); pages.push(pg);
        pg.text(room.name, M + 16, M + 26, 15.5, true, 0);
        y = M + 38;
        pg.fillRect(tx, y, tWidth, HDR, '#EEF3FA');
        HEAD.forEach((h, i) => { pg.strokeRect(colX[i], y, cw[i], HDR, 0.7); pg.text(h, colX[i] + PAD, y + HDR / 2 + 2.7, 7.6, true, 0); });
        y += HDR;
      };
      const close = () => pg.card(M, M, PW - 2 * M, y + 9 - M, 10, 0.9);
      start();
      room.rows.forEach(r => {
        const m = measure(r);
        if (y + m.h > LIMIT && y > M + 38 + HDR) { close(); start(); }
        drawRow(pg, m, y);
        y += m.h;
      });
      close();
    });

    return { blob: assemble(pages, list, title), missing, pages: pages.length };
  }

  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }

  globalThis.FB_pdf = { create, download };
})();
