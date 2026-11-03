// Direct port of the Apps Script extraction rules. Pure functions, no I/O.
const CONFIG = require('./config');

const catByKey = new Map(CONFIG.CATEGORIES.map(c => [c.key, c]));
const catIndex = new Map(CONFIG.CATEGORIES.map((c, i) => [c.key, i]));

function mapColumns(data) {
  const map = { item: -1, carcassMat: -1, shutterMat: -1, shutterFin: -1, carcassFin: -1, qty: -1, startRow: -1 };
  let found = 0;
  for (let i = 0; i < Math.min(20, data.length); i++) {
    if (!data[i]) continue;
    for (let j = 0; j < data[i].length; j++) {
      const val = String(data[i][j]).trim().toUpperCase();
      if (!val) continue;
      if (val === 'ITEM' && map.item === -1) { map.item = j; map.startRow = i + 2; found++; }
      else if (val.includes('CARCASS MATERIAL') && map.carcassMat === -1) { map.carcassMat = j; found++; }
      else if (val.includes('SHUTTER MATERIAL') && map.shutterMat === -1) { map.shutterMat = j; found++; }
      else if (val.includes('SHUTTER FINISH') && map.shutterFin === -1) { map.shutterFin = j; found++; }
      else if (val.includes('CARCASS FINISH') && map.carcassFin === -1) { map.carcassFin = j; found++; }
      else if ((val === 'QTY' || val === 'QTY.' || val === 'QUANTITY') && map.qty === -1) { map.qty = j; found++; }
    }
    if (found >= 6) break;
  }
  return map;
}

function getCategoryKey(name, dict) {
  if (dict[name]) {
    const cat = CONFIG.CATEGORIES.find(c => c.codes.includes(dict[name]));
    if (cat) return cat.key;
  }
  for (const cat of CONFIG.CATEGORIES) {
    if (cat.keywords.some(kw => name.includes(kw))) return cat.key;
  }
  return null;
}

function extractData(data, colMap, db) {
  const rooms = [];
  let currentRoom = null;
  let isBlockedSection = false;
  const accessoryList = db.accessories;
  const cabinetDict = db.cabinets;
  const cell = (row, idx) => (idx >= 0 && row[idx] != null ? String(row[idx]).trim() : '');

  for (let i = colMap.startRow; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const rowStr = row.join(' ').toLowerCase();
    if (CONFIG.END_TRIGGERS.some(t => rowStr.includes(t))) break;

    const itemNameRaw = cell(row, colMap.item);
    const itemNameLower = itemNameRaw.toLowerCase();
    const rowCandidate = itemNameRaw !== '' ? itemNameRaw : String(row[0] == null ? '' : row[0]).trim();
    const candidateLower = rowCandidate.toLowerCase();
    const slNoVal = colMap.item > 0 ? String(row[colMap.item - 1] == null ? '' : row[colMap.item - 1]).trim() : '';
    const hasSlNo = slNoVal !== '' && !isNaN(Number(slNoVal));
    const filledCols = row.filter(c => String(c == null ? '' : c).trim() !== '').length;
    const isWicker = /wicker|whicker/i.test(candidateLower);

    if (!hasSlNo && rowCandidate !== '' && filledCols < 4) {
      const isAccessoryHeader = accessoryList.some(kw => candidateLower.includes(kw)) && !isWicker;
      if (CONFIG.BLOCKER_HEADERS.some(kw => candidateLower.includes(kw)) || isAccessoryHeader) {
        isBlockedSection = true;
        continue;
      }
    }
    if (CONFIG.IGNORED_ITEMS.some(kw => itemNameLower.includes(kw))) continue;
    if (accessoryList.some(kw => itemNameLower.includes(kw)) && !/wicker|whicker/i.test(itemNameLower)) continue;

    if (!hasSlNo && rowCandidate !== '' && filledCols < 4 && isNaN(rowCandidate) && rowCandidate.length > 2 &&
        !candidateLower.includes('spec') && !candidateLower.includes('carcass')) {
      if (candidateLower.includes('room total amount') || !CONFIG.IGNORED_HEADERS.some(kw => candidateLower.includes(kw))) {
        let cleanName = rowCandidate.replace(/\s*-\s*Room\s*total\s*amount.*/i, '').trim();
        cleanName = cleanName.replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.substr(1).toLowerCase());
        if (cleanName !== '' && !CONFIG.IGNORED_HEADERS.includes(cleanName.toLowerCase())) {
          currentRoom = { name: cleanName, categories: [] };
          rooms.push(currentRoom);
          isBlockedSection = false;
          continue;
        }
      }
    }

    if (!isBlockedSection && itemNameRaw !== '' && hasSlNo) {
      let sMat = cell(row, colMap.shutterMat);
      let cMat = cell(row, colMap.carcassMat);
      let sFin = cell(row, colMap.shutterFin);
      let cFin = cell(row, colMap.carcassFin);
      const qty = parseFloat(row[colMap.qty]) || 1;
      let key = getCategoryKey(itemNameLower, cabinetDict);
      const isBothSide = itemNameLower.includes('both side') || itemNameLower.includes('both sides') || (key && key.includes('_both'));
      const catConfig = catByKey.get(key);
      const isCabinetCategory = catConfig ? catConfig.isCabinet : true;

      if (itemNameLower.includes('glass') && isCabinetCategory) {
        if (itemNameLower.includes('wall') || (key && key.includes('wall'))) key = 'wall_glass_unit';
        else if (itemNameLower.includes('tall') || (key && key.includes('tall'))) key = 'tall_glass_unit';
        else key = 'base_glass_unit';
        sMat = 'Glass'; sFin = 'Glass';
      } else if (itemNameLower.includes('open') && isCabinetCategory) {
        if (sMat !== '') key = itemNameLower.includes('wall') ? 'wall' : (itemNameLower.includes('tall') ? 'tall' : 'base');
        else key = itemNameLower.includes('wall') ? 'wall_open_unit' : (itemNameLower.includes('tall') ? 'tall_open_unit' : 'base_open_unit');
      }
      if (isBothSide) {
        const excelFin = cFin || sFin; cFin = excelFin; sFin = excelFin;
        const excelMat = cMat || sMat; cMat = excelMat; sMat = excelMat;
      }

      if (key) {
        if (cMat === '' && sMat === '' && cFin === '' && sFin === '') continue;
        if (!currentRoom) { currentRoom = { name: 'General Area', categories: [] }; rooms.push(currentRoom); }
        if (sMat === '' && !key.includes('glass') && !isBothSide) sMat = cMat;
        if (sFin === '' && !key.includes('glass') && !isBothSide) sFin = cFin;
        const finalCat = catByKey.get(key);
        let displayName = finalCat ? finalCat.name : itemNameRaw;
        if (isBothSide && !displayName.toLowerCase().includes('both side')) displayName += ' (Both Side)';
        currentRoom.categories.push({ key, name: displayName, count: qty, cMat, sMat, cFin, sFin });
      }
    }
  }
  return rooms;
}

function applyIntelligentEndPanels(rooms) {
  const singleSided18 = ['end_panel_18', 'architrave'];
  const singleSided12 = ['end_panel_12'];
  const FILL = 'Fill Manually (there are no cabinets with carcass finish)';
  const mostFrequent = counts => Object.keys(counts).reduce((a, b) => (counts[a] > counts[b] ? a : b), '');

  rooms.forEach(room => {
    const finCounts = {}, matCounts = {};
    room.categories.forEach(cat => {
      const cc = catByKey.get(cat.key);
      if (cc && cc.isCabinet) {
        if (cat.cFin && cat.cFin.trim() !== '') finCounts[cat.cFin] = (finCounts[cat.cFin] || 0) + cat.count;
        if (cat.cMat && cat.cMat.trim() !== '') matCounts[cat.cMat] = (matCounts[cat.cMat] || 0) + cat.count;
      }
    });
    const mostFin = mostFrequent(finCounts);
    const mostMat = mostFrequent(matCounts);

    room.categories.forEach(cat => {
      const isBothSide = cat.key && cat.key.includes('_both');
      if (isBothSide) {
        const f = cat.cFin || cat.sFin; cat.cFin = f; cat.sFin = f;
        const m = cat.cMat || cat.sMat; cat.cMat = m; cat.sMat = m;
      } else if (singleSided12.includes(cat.key)) {
        const origFin = cat.cFin; const origMat = cat.cMat || cat.sMat;
        cat.sFin = origFin;
        cat.cFin = mostFin !== '' ? mostFin : FILL;
        cat.cMat = origMat; cat.sMat = origMat;
      } else if (singleSided18.includes(cat.key)) {
        const origFin = cat.cFin; const origMat = cat.cMat;
        cat.sFin = origFin;
        cat.cFin = mostFin !== '' ? mostFin : FILL;
        cat.sMat = origMat;
        cat.cMat = mostMat !== '' ? mostMat : FILL;
      }
    });
  });
}

function applySmartMerging(rooms) {
  const s = str => String(str || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  rooms.forEach(room => {
    const merged = new Map();
    room.categories.forEach(cat => {
      const k = [s(cat.name), s(cat.cMat), s(cat.sMat), s(cat.cFin), s(cat.sFin)].join('|');
      const ex = merged.get(k);
      if (ex) ex.count += (cat.count || 1);
      else merged.set(k, cat);
    });
    room.categories = [...merged.values()]
      .filter(c => c.count > 0)
      .sort((a, b) => (catIndex.has(a.key) ? catIndex.get(a.key) : 999) - (catIndex.has(b.key) ? catIndex.get(b.key) : 999));
  });
}

function extractProjectName(data) {
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    if (!row) continue;
    for (let j = 0; j < row.length; j++) {
      if (String(row[j]).trim().toLowerCase().includes('project id')) {
        for (let k = j + 1; k < row.length; k++) {
          const name = String(row[k]).trim();
          if (name !== '') return name.replace(/[\/\\\?\*\:\[\]]/g, '').trim().substring(0, 30);
        }
      }
    }
  }
  return 'Extract_' + Date.now();
}

function getLaminateDetails(finish, dict) {
  if (!finish) return { lines: [], image: '' };
  const key = finish.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  if (key.includes('fillmanually') || finish.trim() === 'Glass') return { lines: [finish.trim()], image: '' };

  const res = { lines: [finish.trim()], image: '' };
  const entry = dict[key];
  if (entry) {
    let baseName = finish.replace(/\([^)]*\)/g, '').replace(/\s+/g, ' ').trim();
    if (entry.type) baseName += ' (' + entry.type + ')';
    const lines = [baseName];
    if (entry.code) lines.push(entry.code);
    if (entry.edgeBand) lines.push(entry.edgeBand.toLowerCase().startsWith('edge band') ? entry.edgeBand : 'Edge Band: ' + entry.edgeBand);
    res.lines = lines;

    if (entry.url) {
      let finalUrl = entry.url;
      const urlMatch = entry.url.match(/(https?:\/\/[^\s"'<>]+)/i);
      if (urlMatch && urlMatch[1]) finalUrl = urlMatch[1].trim();
      const driveMatch = finalUrl.match(/(?:id=|\/d\/)([a-zA-Z0-9-_]+)/);
      if (driveMatch && driveMatch[1]) finalUrl = 'https://lh3.googleusercontent.com/d/' + driveMatch[1];
      if (finalUrl.startsWith('http')) res.image = finalUrl;
    }
  }
  return res;
}

function buildReport(rooms, lamDict) {
  const out = [];
  rooms.forEach(room => {
    if (room.categories.length === 0) return;
    out.push({
      name: room.name,
      rows: room.categories.map(c => ({
        product: c.count > 1 ? `${c.name} (Qty: ${c.count})` : c.name,
        carcassMaterial: c.cMat,
        shutterMaterial: c.sMat,
        carcass: getLaminateDetails(c.cFin, lamDict),
        shutter: getLaminateDetails(c.sFin, lamDict)
      }))
    });
  });
  return out;
}

function processMatrix(dataMatrix, db) {
  const colMap = mapColumns(dataMatrix);
  if (colMap.item === -1) throw new Error("Could not find 'ITEM' column. Please check your Template.");
  const rooms = extractData(dataMatrix, colMap, db);
  applyIntelligentEndPanels(rooms);
  applySmartMerging(rooms);
  return { projectName: extractProjectName(dataMatrix), rooms: buildReport(rooms, db.laminates) };
}

module.exports = { processMatrix };
