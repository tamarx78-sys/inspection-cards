// オフライン用: アプリ本体をキャッシュし、キャッシュ優先で返す。
// アプリを更新したら VERSION を上げること。
const VERSION = 'v14';
const CACHE = `inspection-cards-${VERSION}`;
const FILES = [
  './', './index.html', './style.css', './app.js', './xlsx.js', './zip.js', './store.js', './profile.js',
  './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith('inspection-cards-') && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  // 開発中 (localhost) はネットワーク優先にして、編集がすぐ反映されるようにする
  if (location.hostname === 'localhost' || location.hostname === '127.0.0.1') {
    e.respondWith(fetch(e.request).catch(() => caches.match(e.request, { ignoreSearch: true })));
    return;
  }
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request)),
  );
});
