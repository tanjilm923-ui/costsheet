/* GIII Costing — data extractor
   Reads "GIII Costing- DK - Copy.xlsx" and writes costing_data.js (window.COSTING_DATA).
   Run:  node extract_data.js  [path-to-xlsx]
*/
const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx.full.min.js'); // local copy of SheetJS

const XLSX_PATH = process.argv[2] || path.join(__dirname, '..', 'GIII Costing- DK - Copy.xlsx');
const OUT = path.join(__dirname, 'costing_data.js');

function num(v) { const n = parseFloat(v); return isFinite(n) ? n : 0; }
function str(v) { return v === undefined || v === null ? '' : String(v).trim(); }

function findLabel(rows, label, col = 1) {
  for (let r = 0; r < rows.length; r++) {
    if (str(rows[r][col]) === label) return r;
  }
  return -1;
}
function rightValue(rows, r, cols = 12) {
  for (let c = 2; c < cols; c++) { const v = rows[r][c]; if (v !== '' && v !== undefined) return v; }
  return '';
}

function parseSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const rec = {
    styleNo: '', season: '', customer: '', fabMill: '', fabricRef: '', fabricDesc: '', fabWidth: '',
    orderNo: '', mover: '', qty: 0, smv: 0, epm: 0.1,
    transportPct: 0, marginPct: 0, commissionPct: 0,
    efficiency: 60, lineCostRate: 0, items: []
  };

  // ---- header fields (label in col B, value somewhere to the right) ----
  const labelMap = {
    'Fab Mill:': 'fabMill', 'Fabric Ref:': 'fabricRef', 'Fabric Description:': 'fabricDesc',
    'Fab Width:': 'fabWidth', 'Style NO:': 'styleNo', 'Order No:': 'orderNo', 'Order No: ': 'orderNo'
  };
  rows.forEach(r => {
    const b = str(r[1]);
    if (labelMap[b] !== undefined) rec[labelMap[b]] = str(rightValue(rows, rows.indexOf(r)));
  });
  // simpler + robust: scan every cell for the label, value = next non-empty cell to the right
  rows.forEach((r, i) => r.forEach((cell, j) => {
    const t = str(cell);
    if (labelMap[t] !== undefined) {
      for (let k = j + 1; k < Math.min(j + 4, r.length); k++) if (r[k] !== '') { rec[labelMap[t]] = r[k]; break; }
    }
    if (t === 'Qty') { for (let k = j + 1; k < Math.min(j + 3, r.length); k++) if (r[k] !== '') { rec.qty = num(r[k]); break; } }
    if (t === 'Transport') { for (let k = j + 1; k < Math.min(j + 3, r.length); k++) if (r[k] !== '' && typeof r[k] === 'number') { rec.transportPct = num(r[k]); break; } }
    if (t === 'Margin Slippage') { for (let k = j + 1; k < Math.min(j + 3, r.length); k++) if (r[k] !== '' && typeof r[k] === 'number') { rec.marginPct = num(r[k]); break; } }
    if (t === 'Commision' || t === 'Commission') { for (let k = j + 1; k < Math.min(j + 3, r.length); k++) if (r[k] !== '' && typeof r[k] === 'number') { rec.commissionPct = num(r[k]); break; } }
    if (t === 'Efficiency') { for (let k = j + 1; k < Math.min(j + 3, r.length); k++) if (r[k] !== '' && typeof r[k] === 'number') { rec.efficiency = num(r[k]); break; } }
    if (t === 'Line cost' || t === 'Line Cost rate') { for (let k = j + 1; k < Math.min(j + 3, r.length); k++) if (r[k] !== '' && typeof r[k] === 'number') { rec.lineCostRate = 0; break; } }
    if (t === 'SMV') {
      for (let k = j + 1; k < r.length; k++) if (typeof r[k] === 'number' && r[k] < 1000) { /* value sits below */ }
      const below = rows[i + 1];
      if (below) for (let k = j; k < below.length; k++) if (typeof below[k] === 'number' && below[k] > 0) { rec.smv = num(below[k]); break; }
    }
    if (t === 'EPM') {
      const below = rows[i + 1];
      if (below) for (let k = j; k < below.length; k++) if (typeof below[k] === 'number' && below[k] > 0 && below[k] <= 10) { rec.epm = num(below[k]); break; }
    }
  }));

  // ---- item rows ----
  const headerR = findLabel(rows, 'Item');
  const ttlFabR = findLabel(rows, 'TTL fabric Cost');
  const ttlTrimR = findLabel(rows, 'TTL Trims Cost');
  const subTotR = findLabel(rows, 'SUB TOTAL');
  const stopR = subTotR > 0 ? subTotR : rows.length;

  let section = 'Fabric';
  for (let i = headerR + 1; i < stopR; i++) {
    const r = rows[i];
    const name = str(r[1]);
    if (name === 'TTL fabric Cost') { section = 'Trims'; continue; }
    if (name === 'TTL Trims Cost') { section = 'Others'; continue; }
    if (!name) continue;
    if (name === 'CM' || name === 'SUB TOTAL' || name === 'Transport' || name === 'Margin Slippage') continue;
    if (/ebmlishments/i.test(name)) continue;
    rec.items.push({
      name,
      ref: str(r[2]),
      size: str(r[3]),
      unit: str(r[4]),
      unitPx: num(r[5]),
      rating: num(r[6]),
      waste: num(r[7]),
      section
    });
  }
  return rec;
}

const wb = XLSX.read(fs.readFileSync(XLSX_PATH), { type: 'buffer' });
const data = wb.SheetNames
  .filter(n => /^DK-/i.test(n))
  .map(n => parseSheet(wb.Sheets[n]));

fs.writeFileSync(OUT, 'window.COSTING_DATA = ' + JSON.stringify(data, null, 2) + ';\n');
data.forEach(d => console.log(d.styleNo, '| qty', d.qty, '| smv', d.smv, '| epm', d.epm,
  '| eff', d.efficiency, '| rate', d.lineCostRate, '| transp', d.transportPct,
  '| items', d.items.length));
console.log('Wrote', OUT);
