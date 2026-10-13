self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open('planificador-cache-v2').then((cache) => {
      return cache.addAll([
        './index.html',
        './manifest.json',
        './css/styles.css',
        './js/app.js',
        './js/storage.js',
        './js/calendar.js',
        './js/activities.js',
        './js/summary.js',
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