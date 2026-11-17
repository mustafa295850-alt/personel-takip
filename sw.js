// Program dosyalarını saklar ki internet yokken de açılsın. Veriler her zaman sunucudan gelir.
const ONBELLEK = 'personel-takip-v1';
const DOSYALAR = ['./', 'index.html', 'ortak.js', 'manifest.webmanifest', 'simge-180.png', 'simge-192.png', 'simge-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(ONBELLEK).then(c => c.addAll(DOSYALAR)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== ONBELLEK).map(x => caches.delete(x)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  // Önce internetten en yenisi, olmazsa saklanan
  e.respondWith(fetch(e.request).then(r => {
    const kopya = r.clone(); caches.open(ONBELLEK).then(c => c.put(e.request, kopya)); return r;
  }).catch(() => caches.match(e.request).then(r => r || caches.match('index.html'))));
});
