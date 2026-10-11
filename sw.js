// ProMuebles Pedidos: abre rápido aunque no haya señal y muestra los avisos de entrega.
const CACHE = 'pm-app-v1';
const ARCHIVOS = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/badge-96.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// Primero la red (así cada mejora llega sola); si no hay señal, la copia guardada.
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(k => k.put(e.request, c)); return r; })
      .catch(() => caches.match(e.request).then(r => r || caches.match('./index.html')))
  );
});

self.addEventListener('push', e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = { body: e.data && e.data.text() }; }
  e.waitUntil(self.registration.showNotification(d.title || 'ProMuebles', {
    body: d.body || '',
    icon: './icons/icon-192.png',
    badge: './icons/badge-96.png',
    tag: d.uid || undefined,
    data: { uid: d.uid || '' }
  }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const uid = e.notification.data && e.notification.data.uid;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => {
    for (const c of cs) { c.postMessage({ abrir: uid }); return c.focus(); }
    return self.clients.openWindow('./' + (uid ? '?p=' + encodeURIComponent(uid) : ''));
  }));
});
