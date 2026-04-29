const fileInput = document.getElementById('fileInput');
const statusBox = document.getElementById('status');

const PRODUCTS = [
  { out: 'Мол гўшти', keys: ['молгўшти','молгушти','молгушти'] },
  { out: 'Картошка', keys: ['картошка'] },
  { out: 'Ўсимлик ёғи', keys: ['кунгабоқарёғи','кунгабокарёги','пахтаёғи','пахтаёги'] },
  { out: 'Шакар', keys: ['шакар'] },
  { out: '1-навли буғдой уни', keys: ['ун(1-нав)','ун1нав','1навлиунданнон'] },
  { out: 'Гуруч', keys: ['гуруч'] }
];

fileInput.addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    setStatus('Excel o‘qilmoqda...');
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const parsed = parseMatrix(matrix);
    const resultWb = buildResultWorkbook(parsed);
    XLSX.writeFile(resultWb, 'Илова_нархлар_таҳлил.xlsx');
    setStatus('Tayyor Excel fayl yuklandi.', true);
  } catch (err) {
    setStatus('Xatolik: ' + err.message, false, true);
  } finally {
    fileInput.value = '';
  }
});

function setStatus(text, ok = false, error = false) {
  statusBox.textContent = text;
  statusBox.className = 'status' + (ok ? ' ok' : '') + (error ? ' error' : '');
}

function norm(v) {
  return String(v ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, '')
    .replace(/ў/g, 'у').replace(/Ў/g, 'у')
    .replace(/қ/g, 'к').replace(/Қ/g, 'к')
    .replace(/ғ/g, 'г').replace(/Ғ/g, 'г')
    .replace(/ҳ/g, 'х').replace(/Ҳ/g, 'х')
    .toLowerCase();
}

function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  const n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

function round100(n) {
  return Math.round(n / 100) * 100;
}

function pct(max, min) {
  if (!min || !max || min <= 0) return 0;
  return Math.round(((max - min) / min) * 100);
}

function parseMatrix(matrix) {
  let headerRow = -1;
  for (let r = 0; r < Math.min(matrix.length, 30); r++) {
    const rowNorm = matrix[r].map(norm);
    if (rowNorm.some(x => x === 'худуд' || x === 'hudud') && rowNorm.some(x => x === 'туман' || x === 'tuman')) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) throw new Error('“Ҳудуд” ва “Туман” ustunlari topilmadi. Faylda header qatori borligini tekshiring.');

  const headers = matrix[headerRow].map(x => String(x || '').trim());
  const nHeaders = headers.map(norm);
  const regionIdx = nHeaders.findIndex(x => x === 'худуд' || x === 'hudud');
  const districtIdx = nHeaders.findIndex(x => x === 'туман' || x === 'tuman');
  if (regionIdx < 0 || districtIdx < 0) throw new Error('“Ҳудуд” yoki “Туман” ustuni aniqlanmadi.');

  const productCols = PRODUCTS.map(p => {
    const idx = nHeaders.findIndex(h => p.keys.some(k => h.includes(norm(k))));
    if (idx < 0) throw new Error('Mahsulot ustuni topilmadi: ' + p.out);
    return { ...p, idx };
  });

  const rows = [];
  for (let r = headerRow + 1; r < matrix.length; r++) {
    const row = matrix[r];
    const region = String(row[regionIdx] || '').trim();
    const district = String(row[districtIdx] || '').trim();
    if (!region) continue;
    if (norm(region).includes('республика')) continue;
    const item = { region, district, values: {} };
    for (const p of productCols) item.values[p.out] = toNum(row[p.idx]);
    rows.push(item);
  }
  if (!rows.length) throw new Error('Maʼlumot qatorlari topilmadi.');
  return { rows, products: PRODUCTS.map(p => p.out) };
}

function minMax(rows, product, useRegionName = false) {
  let maxRow = null, minRow = null;
  for (const row of rows) {
    const value = row.values[product];
    if (value === null) continue;
    if (!maxRow || value > maxRow.value) maxRow = { row, value };
    if (!minRow || value < minRow.value) minRow = { row, value };
  }
  if (!maxRow || !minRow || maxRow.value <= minRow.value) return null;
  const maxPrice = round100(maxRow.value);
  const minPrice = round100(minRow.value);
  return [product, useRegionName ? maxRow.row.region : maxRow.row.district, maxPrice, useRegionName ? minRow.row.region : minRow.row.district, minPrice, pct(maxPrice, minPrice) + '%'];
}

function buildResultWorkbook(parsed) {
  const regionRows = parsed.rows.filter(r => !r.district);
  const districtRows = parsed.rows.filter(r => r.district);

  const section1 = parsed.products.map(p => minMax(regionRows, p, true)).filter(Boolean);
  const section3 = parsed.products.map(p => minMax(districtRows, p, false)).filter(Boolean);

  const section2 = [];
  for (const product of parsed.products) {
    const groups = new Map();
    for (const row of districtRows) {
      if (!groups.has(row.region)) groups.set(row.region, []);
      groups.get(row.region).push(row);
    }
    let best = null;
    for (const [region, rows] of groups) {
      const mm = minMax(rows, product, false);
      if (!mm) continue;
      const d = parseInt(mm[5], 10);
      if (!best || d > best.diff) best = { region, row: mm, diff: d };
    }
    if (best) {
      best.row[0] = `${product} (${best.region})`;
      section2.push(best.row);
    }
  }

  const wsData = [];
  wsData.push(['', '', '', '', '', '', 'Илова']);
  addSection(wsData, 'Республика бўйича энг юқори ва энг паст нархлари шаклланган ҳудудлар', section1);
  addSection(wsData, 'Бир ҳудудда туманлар кесимида кескин фарқланиш ҳолатлари', section2);
  addSection(wsData, 'Республика бўйича энг қиммат ва энг арзон шаклланган туманлар', section3);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{wch:5},{wch:34},{wch:18},{wch:12},{wch:22},{wch:12},{wch:10}];
  ws['!merges'] = [];
  for (let r = 0; r < wsData.length; r++) {
    if (wsData[r][1] && String(wsData[r][1]).startsWith('Республика') || String(wsData[r][1] || '').startsWith('Бир')) {
      ws['!merges'].push({ s:{r,c:1}, e:{r,c:6} });
    }
  }
  styleSheet(ws, wsData.length);
  XLSX.utils.book_append_sheet(wb, ws, 'Илова');
  return wb;
}

function addSection(wsData, title, rows) {
  wsData.push([]);
  wsData.push(['', title, '', '', '', '', '']);
  wsData.push([]);
  wsData.push(['№', 'Товарлар', 'Энг қиммат', 'Нархи', 'Энг арзон', 'Нархи', 'Фарқи']);
  rows.forEach((r, i) => wsData.push([i + 1, ...r]));
}

function cell(ws, r, c) {
  return ws[XLSX.utils.encode_cell({ r, c })];
}

function styleSheet(ws, totalRows) {
  const border = { style: 'thin', color: { rgb: '777777' } };
  for (let r = 0; r < totalRows; r++) {
    for (let c = 0; c < 7; c++) {
      const x = cell(ws, r, c);
      if (!x) continue;
      x.s = { font: { name: 'Arial', sz: 14 }, alignment: { horizontal: 'center', vertical: 'center' } };
      if (c === 1) x.s.alignment.horizontal = 'left';
      const val = String(x.v || '');
      if (val.startsWith('Республика') || val.startsWith('Бир')) {
        x.s.font = { name: 'Arial', bold: true, sz: 18, color: { rgb: '0070C0' } };
        x.s.alignment.horizontal = 'center';
      }
      if (['№','Товарлар','Энг қиммат','Нархи','Энг арзон','Фарқи'].includes(val)) {
        x.s.fill = { fgColor: { rgb: 'E2F0D9' } };
        x.s.font = { name: 'Arial', bold: true, sz: 14 };
        x.s.border = { top: border, bottom: border, left: border, right: border };
      }
      if (r > 0 && cell(ws, r, 0) && typeof cell(ws, r, 0).v === 'number') {
        x.s.border = { top: border, bottom: border, left: border, right: border };
        if (c === 3) x.s.font = { name: 'Arial', bold: true, sz: 14, color: { rgb: 'FF0000' } };
        if (c === 5) x.s.font = { name: 'Arial', bold: true, sz: 14, color: { rgb: '00A651' } };
        if (c === 6) x.s.font = { name: 'Arial', italic: true, sz: 14 };
        if (c === 3 || c === 5) x.z = '# ##0';
      }
    }
  }
}
