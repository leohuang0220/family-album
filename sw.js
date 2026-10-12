// 只快取 App 外殼；照片不存進快取，避免留在共用裝置上
const C = 'family-album-v5';
const SHELL = ['./', './index.html', './b-icon.svg', './b-icon-192.png', './b-icon-512.png', './b-icon-maskable-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(C).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  if (u.pathname.endsWith('.webmanifest') || u.pathname.endsWith('manifest.json')) return; // 安裝設定一律抓最新
  e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(C).then(c => c.put(e.request, cp)); return r; }).catch(() => caches.match(e.request)));
});
