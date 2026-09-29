#!/usr/bin/env node
/**
 * Rebuild data/indonesia-universities.json from daftar-kampus.xlsx
 * (columns: Nama | Singkatan | NPSN).
 *
 * Usage:
 *   node scripts/fetch-universities.js
 *   node scripts/fetch-universities.js --from ./data/daftar-kampus.xlsx
 *   node scripts/fetch-universities.js --from ~/Downloads/daftar-kampus.xlsx
 */
'use strict';

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const DEFAULT_FROM = path.join(__dirname, '..', 'data', 'daftar-kampus.xlsx');
const OUT = path.join(__dirname, '..', 'data', 'indonesia-universities.json');

function parseArgs(argv) {
  let from = DEFAULT_FROM;
  for (let i = 2; i < argv.length; i++) {
    if (argv[i] === '--from' && argv[i + 1]) from = path.resolve(argv[++i]);
  }
  return { from };
}

function cellText(value) {
  if (value == null) return '';
  if (typeof value === 'object') {
    if (value.richText) return value.richText.map((t) => t.text).join('');
    if (value.text != null) return String(value.text);
    if (value.result != null) return String(value.result);
  }
  return String(value);
}

async function main() {
  const { from } = parseArgs(process.argv);
  if (!fs.existsSync(from)) {
    console.error(`File not found: ${from}`);
    process.exit(1);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(from);
  const ws = wb.getWorksheet('kampus') || wb.worksheets[0];
  if (!ws) {
    console.error('No worksheet found in workbook.');
    process.exit(1);
  }

  const names = [];
  ws.eachRow((row, n) => {
    if (n === 1) return; // header
    const name = cellText(row.getCell(1).value).replace(/\s+/g, ' ').trim();
    if (name.length >= 3) names.push(name);
  });

  const uniq = [...new Set(names)].sort((a, b) => a.localeCompare(b, 'id'));
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(uniq));
  console.log(`Wrote ${uniq.length} names → ${OUT}`);
  console.log(`Source: ${from}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
