// Valley View Sunday School - service worker
// - App files: network first (so updates arrive), cached copy when offline.
// - Tailwind / fonts / QR library: cached after first load so the app still looks right offline.
// - Server calls (Apps Script) and link-preview images are never cached.
const VERSION = 'v4';
const SHELL_CACHE = 'vv-shell-' + VERSION;
const RUNTIME_CACHE = 'vv-runtime-' + VERSION;
const SHELL = [
  './', './index.html', './manifest.json',
  './icon-192.png', './icon-512.png', './icon-maskable-512.png', './apple-touch-icon.png', './favicon-32.png'
];
const CDN_HOSTS = [
  'cdn.tailwindcss.com', 'cdnjs.cloudflare.com',
  'fonts.googleapis.com', 'fonts.gstatic.com', 'fonts.cdnfonts.com'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE).then((cache) =>
      Promise.all(SHELL.map((u) => cache.add(u).catch(() => {})))
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== SHELL_CACHE && k !== RUNTIME_CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function cacheable(res) {
  return res && (res.ok || res.type === 'opaque');
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  const sameOrigin = url.origin === self.location.origin;
  const isCdn = CDN_HOSTS.indexOf(url.hostname) !== -1;
  if (!sameOrigin && !isCdn) return; // Apps Script, images, anything else: straight to the network

  // Pages: network first, fall back to the saved copy
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(SHELL_CACHE).then((c) => c.put('./index.html', copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => caches.match('./index.html').then((r) => r || caches.match('./')))
    );
    return;
  }

  // Everything else we allow: serve the saved copy right away and refresh it in the background
  event.respondWith(
    caches.open(sameOrigin ? SHELL_CACHE : RUNTIME_CACHE).then((cache) =>
      cache.match(req).then((cached) => {
        const network = fetch(req)
          .then((res) => { if (cacheable(res)) cache.put(req, res.clone()).catch(() => {}); return res; })
          .catch(() => cached);
        return cached || network;
      })
    )
  );
});
