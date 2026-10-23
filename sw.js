const CACHE_NAME = 'excel-viewer-v5';
const ASSETS = [
  './',
  './index.html',
  './css/style.css',
  './css/leaflet.css',
  './css/MarkerCluster.css',
  './css/MarkerCluster.Default.css',
  './css/images/layers.png',
  './css/images/layers-2x.png',
  './css/images/marker-icon.png',
  './css/images/marker-icon-2x.png',
  './css/images/marker-shadow.png',
  './js/app.js',
  './js/map.js',
  './js/leaflet.js',
  './js/leaflet.markercluster.js',
  './manifest.json',
  './Data.xlsx',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/eremedium-logo.svg',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS).catch(() => {}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(response => {
      if (response && response.status === 200 && response.type !== 'opaque') {
        const responseClone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(e.request, responseClone));
      }
      return response;
    }).catch(() => cached))
  );
});
