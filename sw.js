/* Service worker: rende l'app installabile e veloce all'avvio.
   I dati NON vengono messi in cache: arrivano sempre dal Google Sheet. */
const CACHE = 'finanze-v2';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const host = new URL(req.url).hostname;
  if (host.endsWith('google.com') || host.endsWith('googleusercontent.com')) return; // API: mai in cache
  if (req.mode === 'navigate') {
    // rete prima (aggiornamenti immediati), cache se offline
    e.respondWith(fetch(req).then(res => {
      const cp = res.clone(); caches.open(CACHE).then(c => c.put('index.html', cp)); return res;
    }).catch(() => caches.match('index.html')));
    return;
  }
  // librerie, font, icone: cache prima, aggiornamento in background
  e.respondWith(caches.match(req).then(hit => {
    const net = fetch(req).then(res => {
      if (res.ok || res.type === 'opaque') { const cp = res.clone(); caches.open(CACHE).then(c => c.put(req, cp)); }
      return res;
    }).catch(() => hit);
    return hit || net;
  }));
});
