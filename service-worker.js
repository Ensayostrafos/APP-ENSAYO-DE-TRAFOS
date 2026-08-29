// Service worker de ENSAYO DE TRAFOS
// Guarda una copia local de la app para que abra sin internet.
// Los ENVÍOS de datos (fetch al Apps Script) NO se cachean acá:
// eso lo maneja la cola de "pendientes" dentro de ensayo_de_trafos.html.

const CACHE_NAME = 'ensayo-trafos-v2';
const ASSETS_TO_CACHE = [
  './ensayo_de_trafos.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = event.request.url;

  // Nunca cacheamos las llamadas al backend (Apps Script): siempre tienen
  // que ir a la red, o fallar para que la app sepa que está offline.
  if (url.includes('script.google.com')) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        // Guardamos en caché las respuestas exitosas de nuestro propio origen o de Leaflet
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => cached);
    })
  );
});
