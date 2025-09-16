// clean_excel.js
// Usage: node clean_excel.js [input.xlsx] [output.xlsx] [--col=1] [--format=wide|long] [--sheet=Sheet1]
// defaults: input.xlsx, output.xlsx, col=1 (first column), format=wide

const ExcelJS = require('exceljs');
const he = require('he');
const fs = require('fs');
const path = require('path');

function parseArgs() {
  const args = process.argv.slice(2);
  const opts = { input: 'input.xlsx', output: 'output.xlsx', col: 1, format: 'wide', sheet: null };
  const positional = [];
  args.forEach(a => {
    if (a.startsWith('--')) {
      const [k,v] = a.slice(2).split('=');
      opts[k] = v === undefined ? true : v;
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

// split top-level by commas but only when the comma is followed by something like "key="
function splitTopLevelFields(str) {
  if (!str) return [];
  return str.split(/,(?=\s*[A-Za-z0-9_\-]+\s*=)/g).map(s => s.trim()).filter(Boolean);
}

function normalizeHtml(value) {
  if (!value) return '';
  let v = value.toString();
  // Fix common bad chars, decode entities
  v = v.replace(/\u00A0/g, ' ')
       .replace(/\u00C2/g, '')   // remove stray Â often seen
       .replace(/&nbsp;/gi, ' ');
  // Keep line breaks for <br> and </p>, then remove other tags
  v = v.replace(/<br\s*\/?>/gi, '\n')
       .replace(/<\/p>/gi, '\n')
       .replace(/<[^>]+>/g, ''); // strip remaining tags
  v = v.replace(/\r/g, '').trim();
  try { v = he.decode(v); } catch(e) {}
  return v;
}

function parseRecord(rawString) {
  const out = {};
  const fields = splitTopLevelFields(rawString);
  for (const f of fields) {
    const idx = f.indexOf('=');
    if (idx === -1) continue;
    let key = f.slice(0, idx).trim();
    let value = f.slice(idx + 1).trim();

    // remove surrounding quotes if present
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }

    // if this is specifications (or contains 'spec'), parse HTML into spec keys
    if (/spec/i.test(key)) {
      const text = normalizeHtml(value);
      const lines = text.split(/\n+/).map(l => l.trim()).filter(Boolean);
      for (const line of lines) {
        if (line.includes(':')) {
          const i = line.indexOf(':');
          const sKey = line.slice(0, i).trim();
          const sVal = line.slice(i+1).trim();
          if (!sKey) continue;
          out[sKey] = out[sKey] ? out[sKey] + ' | ' + sVal : sVal;
        } else {
          // no colon: store under 'spec_note' (or append to an existing spec field)
          out['spec_note'] = out['spec_note'] ? out['spec_note'] + ' | ' + line : line;
        }
      }
    } else {
      // generic key: strip HTML entities and tags
      const cleanVal = normalizeHtml(value);
      out[key] = out[key] ? out[key] + ' | ' + cleanVal : cleanVal;
    }
  }
  return out;
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

  const records = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    // find the first cell in this row that looks like a raw field (contains '=')
    let rawCell = null;
    // prefer configured column
    const preferCell = row.getCell(opts.col);
    if (preferCell && preferCell.value && String(preferCell.value).includes('=')) {
      rawCell = preferCell;
    } else {
      // fallback: scan row for any cell with '='
      for (let i = 1; i <= row.cellCount; i++) {
        const c = row.getCell(i);
        if (c && c.value && String(c.value).includes('=')) {
          rawCell = c;
          break;
        }
      }
    }

    if (!rawCell || !rawCell.value) return;
    let raw = rawCell.value;
    // handle RichText cell types
    if (typeof raw === 'object') {
      if (raw.richText) raw = raw.richText.map(t => t.text).join('');
      else if (raw.text) raw = raw.text;
      else raw = String(raw);
    } else raw = String(raw);

    const parsed = parseRecord(raw);
    parsed.__rowNumber = rowNumber;
    records.push(parsed);
  });

  if (records.length === 0) {
    console.log('No records found.');
    process.exit(0);
  }

  // collect all unique keys for wide output
  const allKeysSet = new Set();
  records.forEach(r => {
    Object.keys(r).forEach(k => {
      if (k === '__rowNumber') return;
      allKeysSet.add(k);
    });
  });
  const allKeys = Array.from(allKeysSet);

  // build output workbook
  const outWb = new ExcelJS.Workbook();
  const outSheet = outWb.addWorksheet('Cleaned');

  if (opts.format === 'long' || opts.format === 'stacked') {
    outSheet.addRow(['Row', 'Attribute', 'Value']);
    for (const rec of records) {
      const rowNum = rec.__rowNumber || '';
      for (const key of allKeys) {
        if (rec[key]) outSheet.addRow([rowNum, key, rec[key]]);
      }
      // blank row between products (optional)
      outSheet.addRow([]);
    }
  } else {
    // wide: header
    outSheet.addRow(['__rowNumber', ...allKeys]);
    for (const rec of records) {
      const row = [rec.__rowNumber || ''];
      for (const key of allKeys) {
        row.push(rec[key] || '');
      }
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
