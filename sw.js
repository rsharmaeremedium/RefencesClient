const CACHE_NAME = 'excel-viewer-v2';
const ASSETS = [
  const ASSETS = [
  '/RefencesClient/',
  '/RefencesClient/index.html',
  '/RefencesClient/css/style.css',
  '/RefencesClient/js/app.js',
  '/RefencesClient/manifest.json',
  '/RefencesClient/Data.xlsx',  
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
    caches.match(e.request).then(cached => cached || fetch(e.request).catch(() => cached))
  );
});
