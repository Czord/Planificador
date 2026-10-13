self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open('planificador-cache-v1').then((cache) => {
      return cache.addAll([
        './index.html',
        './manifest.json',
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
        'https://cdn.jsdelivr.net/npm/chart.js'
      ]);
    })
  );
});

self.addEventListener('fetch', (e) => {
  e.respondWith(
    caches.match(e.request).then((response) => {
      return response || fetch(e.request);
    })
  );
});