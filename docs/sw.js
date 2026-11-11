// Rad Quest service worker: keeps the app shell on the device so it opens instantly, even offline.
// It never stores anything from the server. Bump the version when you change the files below.
const VERSION = 'radquest-shell-3.0.0';
const SHELL = ['index.html', 'config.js', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];
const FONT_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(VERSION).then((cache) => cache.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith('radquest-') && k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Serve from the device first, and quietly fetch a fresh copy for next time.
function staleWhileRevalidate(cacheKey, request) {
  return caches.open(VERSION).then((cache) =>
    cache.match(cacheKey).then((hit) => {
      const fresh = fetch(request).then((res) => {
        if (res && (res.ok || res.type === 'opaque')) cache.put(cacheKey, res.clone());
        return res;
      });
      if (hit) { fresh.catch(() => {}); return hit; }
      return fresh;
    })
  );
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;                       // server calls are POSTs: always live, never stored
  const url = new URL(req.url);

  if (url.origin === self.location.origin) {
    if (req.mode === 'navigate') {
      // Personal links differ only by ?u=CODE, so every one of them is the same page. The code is never stored here.
      event.respondWith(staleWhileRevalidate('index.html', new Request('index.html', { cache: 'no-cache' })));
      return;
    }
    const path = url.pathname.replace(new URL('./', self.location).pathname, '');
    if (SHELL.indexOf(path) >= 0) {
      event.respondWith(staleWhileRevalidate(path, new Request(path, { cache: 'no-cache' })));
    }
    return;
  }
  if (FONT_HOSTS.indexOf(url.hostname) >= 0) {
    event.respondWith(staleWhileRevalidate(req.url, req));
  }
});
