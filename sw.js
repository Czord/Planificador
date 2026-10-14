const CACHE_NAME = 'planificador-cache-v-dynamic';

// Instalación básica del Service Worker
self.addEventListener('install', (event) => {
  self.skipWaiting();
});

// Activación y limpieza de versiones muy antiguas
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Estrategia: "Network First" (Busca primero en la web, si no hay internet usa la caché)
self.addEventListener('fetch', (event) => {
  event.respondWith(
    fetch(event.request)
      .then((networkResponse) => {
        // Si hay internet, clonamos la respuesta fresca y la guardamos en caché automáticamente
        return caches.open(CACHE_NAME).then((cache) => {
          cache.put(event.request, networkResponse.clone());
          return networkResponse;
        });
      })
      .catch(() => {
        // Si el usuario está sin conexión, servimos lo que tengamos guardado
        return caches.match(event.request);
      })
  );
});