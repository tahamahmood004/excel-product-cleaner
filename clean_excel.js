// clean_excel.js
// Usage:
//   node clean_excel.js [input.xlsx] [output.xlsx] [--col=1] [--format=wide|long|sku-attrs] [--sheet=Sheet1]
// New for sku-attrs:
//   --skuCol=sku                  // header (case-insensitive) for SKU column (default: sku)
//   --blobCol=additional_attributes // OPTIONAL header to force which column holds the key=value HTML blob
//   --includeParent=1             // also output ParentSKU
//   --suffixes=BLACK,BLUE,GREEN,RED,WHITE,SILVER,GOLD,GRAY,GREY
//
// Defaults: input.xlsx, output.xlsx, col=1, format=wide

const ExcelJS = require('exceljs');
const he = require('he');
const fs = require('fs');

function parseArgs() {
  const args = process.argv.slice(2);
  const opts = {
    input: 'input.xlsx',
    output: 'output.xlsx',
    col: 1,
    format: 'wide',
    sheet: null,
    skuCol: 'sku',
    blobCol: '',
    includeParent: false,
    suffixes: 'BLACK,BLUE,GREEN,RED,WHITE,SILVER,GOLD,GRAY,GREY',
  };
  const positional = [];
  args.forEach(a => {
    if (a.startsWith('--')) {
      const [k, v] = a.slice(2).split('=');
      if (k === 'includeParent') {
        opts.includeParent = v === undefined ? true : (v === '1' || v === 'true');
      } else {
        opts[k] = v === undefined ? true : v;
      }
    } else {
      positional.push(a);
    }
  });
  if (positional[0]) opts.input = positional[0];
  if (positional[1]) opts.output = positional[1];
  opts.col = Number(opts.col) || 1;
  opts.format = (opts.format || 'wide').toLowerCase();
  return opts;
}

// split by commas only when followed by something like "key="
function splitTopLevelFields(str) {
  if (!str) return [];
  return str.split(/,(?=\s*[A-Za-z0-9_\-]+\s*=)/g).map(s => s.trim()).filter(Boolean);
}

// normalize + strip HTML, keep line breaks for common tags
function normalizeHtml(value) {
  if (!value) return '';
  let v = value.toString();

  // Fix common odd chars and entities (Â, NBSP, etc.)
  v = v.replace(/\u00A0/g, ' ')
       .replace(/\u00C2/g, '')
       .replace(/&nbsp;/gi, ' ');

  // Preserve line breaks for spec-like parsing, then strip tags
  v = v.replace(/<br\s*\/?>/gi, '\n')
       .replace(/<\/p>/gi, '\n')
       .replace(/<\/li>/gi, '\n')
       .replace(/<\/div>/gi, '\n')
       .replace(/<[^>]+>/g, '');

  v = v.replace(/\r/g, '').trim();

  try { v = he.decode(v); } catch (e) {}
  return v;
}

// parse "key=value, key2=value2, ..." and also lines like "CPU: ..." inside values
function parseRecord(rawString) {
  const out = {};
  const fields = splitTopLevelFields(rawString);
  for (const f of fields) {
    const idx = f.indexOf('=');
    if (idx === -1) continue;
    let key = f.slice(0, idx).trim();
    let value = f.slice(idx + 1).trim();

    // drop surrounding quotes
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }

    const clean = normalizeHtml(value);

    // main key=value
    if (key) {
      out[key] = out[key] ? out[key] + ' | ' + clean : clean;
    }

    // sub-attributes inside value (e.g., "CPU: X\nGPU: Y")
    const lines = clean.split(/\n+/).map(s => s.trim()).filter(Boolean);
    for (const line of lines) {
      const cidx = line.indexOf(':');
      if (cidx > 0) {
        const subK = line.slice(0, cidx).trim();
        const subV = line.slice(cidx + 1).trim();
        if (subK) {
          out[subK] = out[subK] ? out[subK] + ' | ' + subV : subV;
        }
      }
    }
  }

  // never promote sku from inside the blob
  delete out['sku']; delete out['SKU']; delete out['Sku'];
  return out;
}

function deriveParentSKU(sku, suffixesArr) {
  if (!sku) return '';
  const U = String(sku).toUpperCase();
  for (const suf of suffixesArr) {
    const S = suf.toUpperCase().trim();
    if (!S) continue;
    if (U.endsWith(S)) {
      return sku.slice(0, sku.length - S.length);
    }
  }
  return '';
}

async function main() {
  const opts = parseArgs();
  if (!fs.existsSync(opts.input)) {
    console.error('Input file not found:', opts.input);
    process.exit(1);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(opts.input);
  const worksheet = opts.sheet ? workbook.getWorksheet(opts.sheet) : workbook.worksheets[0];
  if (!worksheet) {
    console.error('Worksheet not found.');
    process.exit(1);
  }

  // Build header index (case-insensitive)
  const headerRow = worksheet.getRow(1);
  const headerMap = {};
  headerRow.eachCell((cell, colNumber) => {
    if (cell && cell.value !== undefined && cell.value !== null) {
      headerMap[String(cell.value).trim().toLowerCase()] = colNumber;
    }
  });

  const skuColName = String(opts.skuCol || 'sku').toLowerCase();
  const skuColIdx = headerMap[skuColName];
  if (!skuColIdx) {
    console.error(`SKU column "${opts.skuCol}" not found.`);
    process.exit(1);
  }

  const blobColName = String(opts.blobCol || '').toLowerCase();
  const blobColIdx = blobColName ? headerMap[blobColName] : null;

  // Parse rows
  const records = []; // { sku, __rowNumber, kv: {attr: value} }
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const skuValRaw = row.getCell(skuColIdx)?.value;
    const sku = skuValRaw ? String(skuValRaw).trim() : '';
    if (!sku) return;

    // Find the blob cell:
    let blobCell = null;
    if (blobColIdx) {
      const c = row.getCell(blobColIdx);
      if (c && c.value) blobCell = c;
    }
    if (!blobCell) {
      // Fallback: look for first cell (excluding sku col) containing '='
      for (let i = 1; i <= row.cellCount; i++) {
        if (i === skuColIdx) continue;
        const c = row.getCell(i);
        if (c && c.value && String(c.value).includes('=')) {
          blobCell = c;
          break;
        }
      }
    }

    if (!blobCell || !blobCell.value) {
      records.push({ sku, __rowNumber: rowNumber, kv: {} });
      return;
    }

    let raw = blobCell.value;
    if (typeof raw === 'object') {
      if (raw.richText) raw = raw.richText.map(t => t.text).join('');
      else if (raw.text) raw = raw.text;
      else raw = String(raw);
    } else raw = String(raw);

    const kv = parseRecord(raw);
    records.push({ sku, __rowNumber: rowNumber, kv });
  });

  if (records.length === 0) {
    console.log('No records found.');
    process.exit(0);
  }

  if (opts.format === 'sku-attrs') {
    const outWb = new ExcelJS.Workbook();
    const outSheet = outWb.addWorksheet('SKU_Attributes');

    const suffixesArr = String(opts.suffixes || '').split(',').map(s => s.trim()).filter(Boolean);
    const includeParent = !!opts.includeParent;

    const header = includeParent ? ['SKU', 'Attribute', 'Value', 'ParentSKU'] : ['SKU', 'Attribute', 'Value'];
    outSheet.addRow(header);

    for (const rec of records) {
      const entries = Object.entries(rec.kv).filter(([k, v]) => v && String(v).trim() !== '');
      if (entries.length === 0) {
        const parent = includeParent ? deriveParentSKU(rec.sku, suffixesArr) : '';
        outSheet.addRow(includeParent ? [rec.sku, '', '', parent] : [rec.sku, '', '']);
        continue;
      }
      const parent = includeParent ? deriveParentSKU(rec.sku, suffixesArr) : '';
      for (const [attr, val] of entries) {
        outSheet.addRow(includeParent ? [rec.sku, attr, val, parent] : [rec.sku, attr, val]);
      }
    }

    await outWb.xlsx.writeFile(opts.output);
    console.log('✅ SKU attributes file written to', opts.output);
    console.log(`Processed ${records.length} row(s). Format=sku-attrs`);
    return;
  }

  // ===== existing outputs (wide/long) preserved =====
  const allKeysSet = new Set();
  records.forEach(r => Object.keys(r.kv || {}).forEach(k => allKeysSet.add(k)));
  const allKeys = Array.from(allKeysSet);

  const outWb = new ExcelJS.Workbook();
  const outSheet = outWb.addWorksheet('Cleaned');

  if (opts.format === 'long' || opts.format === 'stacked') {
    outSheet.addRow(['SKU', 'Row', 'Attribute', 'Value']);
    for (const rec of records) {
      for (const key of allKeys) {
        const val = (rec.kv || {})[key];
        if (val) outSheet.addRow([rec.sku, rec.__rowNumber || '', key, val]);
      }
      outSheet.addRow([]);
    }
  } else {
    outSheet.addRow(['SKU', '__rowNumber', ...allKeys]);
    for (const rec of records) {
      const row = [rec.sku, rec.__rowNumber || ''];
      for (const key of allKeys) row.push((rec.kv || {})[key] || '');
      outSheet.addRow(row);
    }
  }

  await outWb.xlsx.writeFile(opts.output);
  console.log('✅ Cleaned file written to', opts.output);
  console.log(`Processed ${records.length} record(s). Format=${opts.format}`);
}

main().catch(err => {
  console.error('Error:', err);
});
