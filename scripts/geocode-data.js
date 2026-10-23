#!/usr/bin/env node
/*
 * Geocodes every unique address in Data.xlsx with OpenStreetMap Nominatim and
 * writes Latitude / Longitude columns back into the sheet, so the Map tab can
 * plot exact locations straight from the file — no lookups at runtime, works
 * offline once the map tiles have been seen.
 *
 *   npm i xlsx
 *   node scripts/geocode-data.js
 *
 * - Polite: 1 request / second, Nominatim usage-policy compliant.
 * - Incremental: rows that already have coordinates are skipped, so re-run it
 *   after adding new clients and only the new addresses get geocoded.
 * - Cascade: full address -> drop the street line -> area + city + pin ->
 *   pin + city -> city centre, so almost every row ends up with something.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');

let XLSX;
try { XLSX = require('xlsx'); }
catch { console.error('Install the reader first:  npm i xlsx'); process.exit(1); }

const XLSX_FILE = path.join(__dirname, '..', 'Data.xlsx');
const CACHE_FILE = path.join(os.tmpdir(), 'erem-geocode-cache.json');
const NOMINATIM = 'https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=in';
const UA = 'EREMEDIUM-ReferencesClient geocoder (github.com/rsharmaeremedium/RefencesClient)';
const GAP = 1100;                       // ms between requests — Nominatim allows 1/s
const HEADERS = ['Latitude', 'Longitude'];

const norm = (s) => String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let cache = {};
try { cache = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8')); } catch {}

let last = 0;
async function ask(q) {
  for (let attempt = 0; ; attempt++) {
    const wait = GAP - (Date.now() - last);
    if (wait > 0) await sleep(wait);
    last = Date.now();
    try {
      const res = await fetch(NOMINATIM + '&q=' + encodeURIComponent(q),
        { headers: { 'User-Agent': UA, 'Accept-Language': 'en' } });
      if (res.status === 429) {                       // rate limited — back off
        await sleep(5000 * (attempt + 1));
        if (attempt < 3) continue;
        return null;
      }
      if (!res.ok) return null;
      const json = await res.json();
      if (!json.length) return null;
      return { lat: +(+json[0].lat).toFixed(6), lng: +(+json[0].lon).toFixed(6), label: json[0].display_name };
    } catch (e) {
      if (attempt >= 2) { console.error('  ! ' + q + ' → ' + e.message); return null; }
      await sleep(2000 * (attempt + 1));
    }
  }
}

// Query variants, most specific first. Long addresses that repeat the city and
// pin code return nothing on Nominatim, so shorter variants come after.
function variants({ address, area, city, pin }) {
  const out = [];
  const push = (q) => { q = q.split(',').map((s) => s.trim()).filter(Boolean).join(', ');
                         if (q && !out.includes(q)) out.push(q); };
  push(address || '');
  if (address) {
    push([address, 'India'].join(', '));
    const segs = address.split(',').map((s) => s.trim()).filter(Boolean);
    if (segs.length > 1) push(segs.slice(1).join(', '));   // drop the first line
  }
  push([area, city, pin].filter(Boolean).join(', '));
  push([pin, city].filter(Boolean).join(', '));
  push(city || '');
  return out;
}

async function locate(row) {
  const key = norm([row.Address, row.Area, row.City, row['Pin Code']].filter(Boolean).join('|'));
  if (cache[key] !== undefined) return cache[key];
  let hit = null;
  for (const q of variants({ address: row.Address, area: row.Area, city: row.City, pin: row['Pin Code'] })) {
    hit = await ask(q);
    if (hit) break;
  }
  cache[key] = hit;                                  // null = tried and failed
  fs.writeFileSync(CACHE_FILE, JSON.stringify(cache));
  return hit;
}

(async () => {
  const wb = XLSX.readFile(XLSX_FILE);
  const sheetName = wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });

  // only rows that still miss coordinates
  const todo = [];
  rows.forEach((r, i) => {
    if (r[HEADERS[0]] !== '' && r[HEADERS[0]] != null) return;
    if (!(r.Address || r.City)) return;             // nothing to search
    todo.push(i);
  });
  const unique = new Set(todo.map((i) =>
    norm([rows[i].Address, rows[i].Area, rows[i].City, rows[i]['Pin Code']].filter(Boolean).join('|'))));
  console.log(`${rows.length} rows · ${todo.length} without coordinates · ${unique.size} unique places to look up`);

  let n = 0, ok = 0;
  for (const i of todo) {
    n++;
    const hit = await locate(rows[i]);
    if (hit) { rows[i][HEADERS[0]] = hit.lat; rows[i][HEADERS[1]] = hit.lng; ok++; }
    if (n % 25 === 0 || n === todo.length) {
      console.log(`  ${n}/${todo.length} looked up · ${ok} found · cache ${Object.keys(cache).length}`);
    }
  }

  // rebuild the sheet with the two coordinate columns appended
  const header = [...Object.keys(rows[0])];
  for (const h of HEADERS) if (!header.includes(h)) header.push(h);
  const ws = XLSX.utils.json_to_sheet(rows, { header });
  wb.Sheets[sheetName] = ws;
  const tmp = XLSX_FILE + '.tmp';
  XLSX.writeFile(wb, tmp, { bookType: 'xlsx' });
  fs.renameSync(tmp, XLSX_FILE);
  console.log(`\nDone — coordinates written to ${sheetName}! (${ok}/${todo.length} newly located)`);
  console.log(`Cache kept at ${CACHE_FILE} — re-runs skip everything already solved.`);
})();
