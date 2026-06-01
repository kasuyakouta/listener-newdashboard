// ================================================================
//  Service Worker — 面談リスナー PWA
// ================================================================

const CACHE_NAME = 'listener-v1';
const STATIC_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-180.png',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      cache.addAll(STATIC_ASSETS).catch(() => {})
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(JSON.stringify({ status: 'error', message: 'オフラインです' }), {
          headers: { 'Content-Type': 'application/json' }
        })
      )
    );
    return;
  }
  event.respondWith(
    caches.match(event.request).then(cached => {
      const networkFetch = fetch(event.request).then(res => {
        if (res.ok) {
          caches.open(CACHE_NAME).then(c => c.put(event.request, res.clone()));
        }
        return res;
      }).catch(() => cached);
      return cached || networkFetch;
    })
  );
});
