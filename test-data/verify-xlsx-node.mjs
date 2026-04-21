/**
 * Parse roster-valid.xlsx with the same xlsx@0.18.5 build as index.html.
 * Run: npm install && node verify-xlsx-node.mjs
 */
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';
import XLSX from 'xlsx';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const buf = fs.readFileSync(path.join(__dirname, 'roster-valid.xlsx'));

const wb = XLSX.read(buf, { type: 'buffer' });
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
if (rows.length !== 7) {
  console.error('Expected 7 data rows, got', rows.length);
  process.exit(1);
}
const keys = Object.keys(rows[0]);
const norm = (k) => String(k ?? '').trim().toLowerCase();
const findCol = (want) => keys.find((k) => norm(k) === want);
if (!findCol('name') || !findCol('type') || !findCol('sector')) {
  console.error('Missing columns', keys);
  process.exit(1);
}
const first = rows[0];
console.log('OK SheetJS:', rows.length, 'rows; first row', findCol('name'), '=', first[findCol('name')]);
