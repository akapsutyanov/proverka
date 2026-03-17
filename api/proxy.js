export const config = {
  api: { bodyParser: { sizeLimit: '20mb' } },
  maxDuration: 60,
};

// Minimal XLSX parser using only Node built-ins (xlsx is a zip)
async function parseXlsx(buffer) {
  const AdmZip = require('adm-zip');
  const { DOMParser } = require('@xmldom/xmldom');

  const zip = new AdmZip(buffer);

  // Read shared strings
  const ssXml = zip.getEntry('xl/sharedStrings.xml');
  const strings = [];
  if (ssXml) {
    const doc = new DOMParser().parseFromString(ssXml.getData().toString('utf8'), 'text/xml');
    const sis = doc.getElementsByTagName('si');
    for (let i = 0; i < sis.length; i++) {
      const ts = sis[i].getElementsByTagName('t');
      let val = '';
      for (let j = 0; j < ts.length; j++) val += ts[j].textContent || '';
      strings.push(val.trim().replace(/\s+/g, ' '));
    }
  }

  // Read workbook to get sheet names
  const wbXml = zip.getEntry('xl/workbook.xml');
  const wbDoc = new DOMParser().parseFromString(wbXml.getData().toString('utf8'), 'text/xml');
  const sheetEls = wbDoc.getElementsByTagName('sheet');
  const sheetMap = {};
  for (let i = 0; i < sheetEls.length; i++) {
    const name = sheetEls[i].getAttribute('name');
    const rid  = sheetEls[i].getAttribute('r:id');
    sheetMap[rid] = name;
  }

  // Read relationships to get sheet file paths
  const relsXml = zip.getEntry('xl/_rels/workbook.xml.rels');
  const relsDoc = new DOMParser().parseFromString(relsXml.getData().toString('utf8'), 'text/xml');
  const rels = relsDoc.getElementsByTagName('Relationship');
  const ridToPath = {};
  for (let i = 0; i < rels.length; i++) {
    ridToPath[rels[i].getAttribute('Id')] = 'xl/' + rels[i].getAttribute('Target');
  }

  const SKIP = new Set(['Доставка','Упаковка','Цвета','Данные']);
  const items = [];

  function cellVal(c) {
    if (!c) return null;
    const t = c.getAttribute('t');
    const vEl = c.getElementsByTagName('v')[0];
    if (!vEl) return null;
    const v = vEl.textContent;
    if (t === 's') return strings[parseInt(v)] || '';
    const n = parseFloat(v);
    return isNaN(n) ? v : n;
  }

  function colNum(ref) {
    const letters = ref.replace(/[0-9]/g, '');
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64;
    return n;
  }

  for (const [rid, shName] of Object.entries(sheetMap)) {
    if (SKIP.has(shName)) continue;
    const path = ridToPath[rid];
    if (!path) continue;
    const entry = zip.getEntry(path.replace('xl/','xl/'));
    if (!entry) continue;

    const doc = new DOMParser().parseFromString(entry.getData().toString('utf8'), 'text/xml');
    const rowEls = doc.getElementsByTagName('row');

    // Build rows as array of {col: value}
    const rowData = [];
    for (let ri = 0; ri < rowEls.length; ri++) {
      const cells = {};
      const cs = rowEls[ri].getElementsByTagName('c');
      for (let ci = 0; ci < cs.length; ci++) {
        const ref = cs[ci].getAttribute('r') || '';
        const col = colNum(ref);
        cells[col] = cellVal(cs[ci]);
      }
      rowData.push(cells);
    }

    // Find header row (col 1 = 'Артикул')
    let hdrIdx = -1;
    for (let i = 0; i < Math.min(rowData.length, 10); i++) {
      if (rowData[i][1] === 'Артикул') { hdrIdx = i; break; }
    }
    if (hdrIdx < 0) continue;

    const hdr = rowData[hdrIdx];
    let colD = null, colPl = null, colNt = null;
    for (const [col, val] of Object.entries(hdr)) {
      if (typeof val === 'string') {
        if (val.includes('Дилерская')) colD = parseInt(col);
        if (val.includes('Места'))     colPl = parseInt(col);
        if (val.includes('Примечание')) colNt = parseInt(col);
      }
    }
    if (!colD) continue;

    let curName = null;
    for (let ri = hdrIdx + 1; ri < rowData.length; ri++) {
      const row = rowData[ri];
      const name = typeof row[2] === 'string' ? row[2] : null;
      if (name) curName = name;
      if (!curName || !row[1]) continue;
      const art = String(row[1]).trim();
      if (!art) continue;
      const dealer = parseFloat(row[colD]) || 0;
      if (!dealer) continue;

      const prices = [];
      for (let k = 0; k < 8; k++) {
        const v = parseFloat(row[colD + k]);
        prices.push(isNaN(v) ? null : Math.round(v));
      }

      items.push({
        sheet: shName,
        art,
        name: curName,
        places: colPl && row[colPl] ? String(row[colPl]).trim() : '',
        note:   colNt && row[colNt] ? String(row[colNt]).trim().replace(/\s+/g,' ') : '',
        dealer: prices[0], opt_large: prices[1], opt_mid: prices[2],
        opt_small: prices[3], disc20: prices[4], disc15: prices[5],
        disc10: prices[6], retail: prices[7],
      });
    }
  }

  return items;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();

  try {
    // 1. Fetch price.xlsx from GitHub repo
    const xlsxResp = await fetch(
      'https://raw.githubusercontent.com/akapsutyanov/proverka/main/price.xlsx'
    );
    if (!xlsxResp.ok) throw new Error('Cannot load price.xlsx from GitHub: ' + xlsxResp.status);
    const xlsxBuf = Buffer.from(await xlsxResp.arrayBuffer());

    // 2. Parse all sheets
    const priceItems = await parseXlsx(xlsxBuf);

    // 3. Build compact price summary for Claude
    const priceText = priceItems
      .map(p => [p.art, p.sheet, p.name + (p.places ? ' ' + p.places : ''), p.note, 'Опт:' + (p.opt_mid || '-')].join(' | '))
      .join('\n');

    // 4. Inject into first text block of request
    const body = req.body;
    if (Array.isArray(body.messages?.[0]?.content)) {
      const first = body.messages[0].content.find(b => b.type === 'text');
      if (first) {
        first.text += `\n\nПОЛНЫЙ ПРАЙС (${priceItems.length} позиций):\n` + priceText;
      }
    }

    // 5. Forward to Anthropic
    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'pdfs-2024-09-25',
      },
      body: JSON.stringify(body),
    });

    const data = await resp.json();
    // Return price items alongside so HTML can build KP
    data._price = priceItems;

    res.setHeader('Content-Type', 'application/json');
    return res.status(resp.status).json(data);

  } catch (err) {
    return res.status(500).json({ error: { message: err.message } });
  }
}
