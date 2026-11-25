// Service worker : met en cache l'interface (hors ligne) ; les données passent toujours par le réseau.
const CACHE = 'formation-v4';
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== CACHE).map(x => caches.delete(x)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET' || new URL(r.url).origin !== location.origin) return; // l'API Apps Script n'est jamais mise en cache
  e.respondWith(fetch(r).then(res => { const cp = res.clone(); caches.open(CACHE).then(c => c.put(r, cp)); return res; }).catch(() => caches.match(r).then(m => m || caches.match('./index.html'))));
});

// Clic sur une notification : ouvre (ou ramène au premier plan) l'application sur le bon cours.
self.addEventListener('notificationclick', e => {
  e.notification.close(); const d = e.notification.data || {};
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => {
    if (cs.length) { cs[0].focus(); cs[0].postMessage({ open: d }); return; }
    return self.clients.openWindow('./index.html' + (d.cid ? '#' + encodeURIComponent(JSON.stringify(d)) : ''));
  }));
});
