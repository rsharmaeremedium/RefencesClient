'use strict';

// ── Map view ───────────────────────────────────────────────
// Reads the Latitude/Longitude columns that scripts/geocode-data.js writes
// into Data.xlsx, groups the rows into one pin per client (same name +
// location = one pin, products combined) and draws them with Leaflet.

let deps = null;             // { escHtml, showToast }
let map = null;
let cluster = null;
let markers = new Map();     // pin key -> L.marker
let pins = [];               // pins for the last render
let lastSignature = '';

const INDIA = [20.59, 78.96];

const nk = (s) => String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
const esc = (s) => (deps ? deps.escHtml(String(s ?? '')) : String(s ?? ''));

// One row per product in the sheet; a doctor listed for three products at one
// address becomes a single pin carrying all three.
function pinKey(row) {
  const loc = row.Address || [row.Area, row.City, row['Pin Code']].filter(Boolean).join(' ');
  return nk(row['Deal Name']) + '||' + nk(loc);
}

function buildPins(rows) {
  const byKey = new Map();
  for (const row of rows) {
    const key = pinKey(row);
    let pin = byKey.get(key);
    if (!pin) {
      pin = {
        key,
        name: String(row['Deal Name'] ?? '').trim(),
        address: String(row.Address ?? '').trim(),
        area: String(row.Area ?? '').trim(),
        city: String(row.City ?? '').trim(),
        pinCode: String(row['Pin Code'] ?? '').trim(),
        products: [], specialities: [], owners: [],
        screens: 0, deals: 0,
        lat: null, lng: null,
      };
      byKey.set(key, pin);
    }
    const add = (list, v) => { v = String(v ?? '').trim(); if (v && !list.includes(v)) list.push(v); };
    // the sheet renamed "Product Category" → "Product Purchased" — accept both.
    // Some cells hold several products ("MedComm; MedXplain; Medio"); split them
    // so each product keeps its own fixed colour in chips/legend.
    String(row['Product Purchased'] ?? row['Product Category'] ?? '')
      .split(/[;,]/).forEach(v => add(pin.products, v));
    add(pin.specialities, row.Speciality);
    add(pin.owners, row['Deal owner']);
    const n = parseFloat(String(row['Number of Screens'] ?? '').replace(/,/g, ''));
    if (!isNaN(n)) pin.screens += n;
    pin.deals++;
    if (pin.lat == null) {
      const lat = parseFloat(row.Latitude), lng = parseFloat(row.Longitude);
      if (isFinite(lat) && isFinite(lng)) { pin.lat = lat; pin.lng = lng; }
    }
  }
  return [...byKey.values()];
}

// Fixed colour per product so a product always looks the same on the map,
// in the legend and in popups. Unknown products fall back to the palette.
const PRODUCT_COLORS = {
  'MedComm': '#1dc7d0',     // teal
  'MedXplain': '#ff6b82',   // coral
  'Medio': '#ffc65c',       // amber
  'MedEmbed': '#9d7bff',    // violet
};
const PALETTE = ['#63d471', '#49ecff', '#ff9f4d', '#f472b6', '#a3e635', '#60a5fa', '#e879f9', '#facc15'];
function colorOf(product) {
  const p = String(product ?? '').trim();
  if (PRODUCT_COLORS[p]) return PRODUCT_COLORS[p];
  let h = 0;
  for (let i = 0; i < p.length; i++) h = (h * 31 + p.charCodeAt(i)) >>> 0;
  return PALETTE[h % PALETTE.length];
}

function ensureMap() {
  if (map) return;
  // zoom control bottom-left: the floating toolbar covers the top of the screen
  map = L.map('map', { zoomControl: false, worldCopyJump: true }).setView(INDIA, 5);
  L.control.zoom({ position: 'bottomleft' }).addTo(map);
  L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
  }).addTo(map);
  cluster = L.markerClusterGroup({ chunkedLoading: true, maxClusterRadius: 45, spiderfyOnMaxZoom: true });
  map.addLayer(cluster);
}

function markerFor(pin) {
  const color = colorOf(pin.products[0]);
  const icon = L.divIcon({
    className: 'pin-icon',
    html: `<span class="pin-dot" style="background:${color}"></span>`,
    iconSize: [15, 15],
    iconAnchor: [8, 8],
    popupAnchor: [0, -8],
  });
  const m = L.marker([pin.lat, pin.lng], { icon, title: pin.name });
  m.bindPopup(() => popupHtml(pin), { maxWidth: 300 });
  return m;
}

function popupHtml(pin) {
  const addr = [pin.address, pin.area, pin.city, pin.pinCode].filter(Boolean).join(', ');
  return `<div class="pin-pop">
    <h4>${esc(pin.name) || '(unnamed)'}</h4>
    ${pin.products.length ? `<div class="pop-chips">${pin.products.map((p) =>
      `<span class="pop-chip" style="border-color:${colorOf(p)}">${esc(p)}</span>`).join('')}</div>` : ''}
    ${pin.deals > 1 ? `<div class="pop-note">${pin.deals} deals combined on this pin</div>` : ''}
    ${pin.screens ? `<div class="pop-row"><b>Screens:</b> ${pin.screens}</div>` : ''}
    ${pin.specialities.length ? `<div class="pop-row"><b>Speciality:</b> ${esc(pin.specialities.join(', '))}</div>` : ''}
    ${pin.owners.length ? `<div class="pop-row"><b>Owner:</b> ${esc(pin.owners.join(', '))}</div>` : ''}
    <div class="pop-addr">${esc(addr) || '—'}</div>
  </div>`;
}

function drawMarkers() {
  const keep = new Set();
  for (const pin of pins) {
    if (pin.lat == null) continue;
    keep.add(pin.key);
    let m = markers.get(pin.key);
    if (m) m.setLatLng([pin.lat, pin.lng]);
    else { m = markerFor(pin); markers.set(pin.key, m); cluster.addLayer(m); }
  }
  for (const [key, m] of markers) {
    if (!keep.has(key)) { cluster.removeLayer(m); markers.delete(key); }
  }
}

function updateStatus() {
  const el = document.getElementById('map-status');
  if (!el) return;
  const placed = pins.filter((p) => p.lat != null).length;
  const missing = pins.length - placed;
  el.innerHTML = `<b>${pins.length.toLocaleString()}</b> pins` +
    (missing ? ` · <span class="warn">${missing.toLocaleString()} without coordinates</span>` : ` · all located`);
}

function updateLegend() {
  const el = document.getElementById('map-legend');
  if (!el) return;
  const seen = new Map();
  for (const p of pins) for (const prod of p.products) if (!seen.has(prod)) seen.set(prod, colorOf(prod));
  el.style.display = seen.size ? '' : 'none';
  el.innerHTML = [...seen.entries()].map(([prod, color]) =>
    `<div><span class="dot" style="background:${color}"></span>${esc(prod)}</div>`).join('');
}

function fitAll() {
  const pts = pins.filter((p) => p.lat != null).map((p) => [p.lat, p.lng]);
  if (!pts.length) { map.setView(INDIA, 5); return; }
  map.fitBounds(L.latLngBounds(pts).pad(0.1), { maxZoom: 14 });
}

function signature(rows) {
  return rows.length + '|' + (rows[0]?.['Deal Name'] ?? '') + '|' + (rows[rows.length - 1]?.['Deal Name'] ?? '');
}

// ── exports (one entry per pin — same grouping as the map) ──
function exportRows() {
  return pins.filter((p) => p.lat != null);
}

function download(name, text, type) {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function exportKml() {
  const rows = exportRows();
  if (!rows.length) { deps.showToast('No located pins to export'); return; }
  const placemarks = rows.map((p) => {
    const addr = [p.address, p.area, p.city, p.pinCode].filter(Boolean).join(', ');
    const desc = [
      p.products.length ? 'Products: ' + p.products.join(', ') : '',
      p.screens ? 'Screens: ' + p.screens : '',
      p.specialities.length ? 'Speciality: ' + p.specialities.join(', ') : '',
      p.owners.length ? 'Owner: ' + p.owners.join(', ') : '',
      addr,
      p.deals > 1 ? p.deals + ' deals combined' : '',
    ].filter(Boolean).join('\n');
    return `  <Placemark>\n    <name>${esc(p.name)}</name>\n    <description>${esc(desc)}</description>\n    <Point><coordinates>${p.lng},${p.lat},0</coordinates></Point>\n  </Placemark>`;
  }).join('\n');
  const kml = `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2">\n<Document>\n  <name>EREMEDIUM clients</name>\n${placemarks}\n</Document>\n</kml>`;
  download(`clients_map_${Date.now()}.kml`, kml, 'application/vnd.google-earth.kml+xml');
  deps.showToast(`🗺 KML ready (${rows.length} pins) — import into Google My Maps`);
}

function exportCsv() {
  const rows = exportRows();
  if (!rows.length) { deps.showToast('No located pins to export'); return; }
  const q = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  const head = ['Name', 'Products', 'Total screens', 'Speciality', 'Owner', 'City', 'Area', 'Address', 'Pin Code', 'Latitude', 'Longitude', 'Deals combined'];
  const body = rows.map((p) => [p.name, p.products.join(' | '), p.screens, p.specialities.join(' | '),
    p.owners.join(' | '), p.city, p.area, p.address, p.pinCode, p.lat, p.lng, p.deals].map(q).join(','));
  download(`clients_map_${Date.now()}.csv`, [head.map(q).join(','), ...body].join('\r\n'), 'text/csv;charset=utf-8');
  deps.showToast(`📄 CSV ready (${rows.length} pins)`);
}

// ── API used by app.js ─────────────────────────────────────
export function initMap(d) {
  deps = d;
  const byId = (id) => document.getElementById(id);
  byId('map-fit').addEventListener('click', () => { fitAll(); });
  byId('map-kml').addEventListener('click', exportKml);
  byId('map-csv').addEventListener('click', exportCsv);
}

export function renderMap(rows) {
  ensureMap();
  pins = buildPins(rows);
  drawMarkers();
  updateStatus();
  updateLegend();
  // Leaflet cannot measure a hidden container — resize once shown, and follow
  // the data when the filter changes the pin set.
  const sig = signature(rows);
  setTimeout(() => {
    map.invalidateSize();
    if (sig !== lastSignature) { lastSignature = sig; fitAll(); }
  }, 60);
}
