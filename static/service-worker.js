const VERSION = 'swiftfin-pwa-2';
const STATIC_CACHE = `static-${VERSION}`;
const RUNTIME_CACHE = `runtime-${VERSION}`;

const PRECACHE = [
  '/offline',
  '/static/styles-pwa.css',
  '/static/pwa-register.js',
  '/static/pwa-core.js'
];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(STATIC_CACHE);
    for (const url of PRECACHE) {
      try {
        await cache.add(url);
      } catch {
        // best-effort precache; don't block install if one asset is unavailable
      }
    }
    self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.map(k => (k !== STATIC_CACHE && k !== RUNTIME_CACHE) ? caches.delete(k) : null));
    await self.clients.claim();
  })());
});

// Only handle GET; POST sync is handled by JS queue on pages
self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  // Navigation: network-first, fallback to cache, then offline page
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        const fresh = await fetch(req);
        const cache = await caches.open(RUNTIME_CACHE);
        cache.put(req, fresh.clone());
        return fresh;
      } catch {
        const cached = await caches.match(req);
        return cached || caches.match('/offline');
      }
    })());
    return;
  }

  // Static: cache-first
  const url = new URL(req.url);
  if (url.pathname.startsWith('/static/')) {
    event.respondWith((async () => {
      const cached = await caches.match(req);
      if (cached) return cached;
      const fresh = await fetch(req);
      const cache = await caches.open(RUNTIME_CACHE);
      cache.put(req, fresh.clone());
      return fresh;
    })());
    return;
  }

  // Default: stale-while-revalidate
  event.respondWith((async () => {
    const cache = await caches.open(RUNTIME_CACHE);
    const cached = await cache.match(req);

    const network = fetch(req).then((fresh) => {
      cache.put(req, fresh.clone());
      return fresh;
    }).catch(() => null);

    return cached || (await network) || new Response('Offline', { status: 503 });
  })());
});
