// 只快取 App 外殼；照片不存進快取，避免留在共用裝置上
const C = 'baby-album-v3';
const SHELL = ['./', './index.html', './app.webmanifest', './c-favicon.png', './c-apple-touch-icon.png', './c-icon-192.png', './c-icon-512.png', './c-icon-maskable-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(C).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C && !k.startsWith('album-media')).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const req = e.request, u = new URL(req.url);
  if (req.method !== 'GET' || u.origin !== location.origin) return;
  if (req.mode === 'navigate' || u.pathname.endsWith('.webmanifest')) {
    e.respondWith(fetch(req).then(r => { const c = r.clone(); caches.open(C).then(x => x.put(req.mode === 'navigate' ? './index.html' : req, c)); return r; })
      .catch(() => caches.match(req.mode === 'navigate' ? './index.html' : req)));
    return;
  }
  e.respondWith(fetch(req).then(r => { const c = r.clone(); caches.open(C).then(x => x.put(req, c)); return r; }).catch(() => caches.match(req)));
});
