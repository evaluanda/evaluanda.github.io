// Service worker mínimo del lanzador instalable (PWA). La app real vive en
// Apps Script (script.google.com) y no se cachea: aquí solo se guarda app.html
// para que el icono instalado abra algo aunque no haya red.
var CACHE = 'evaluanda-lanzador-v1';
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.add('/app.html'); }));
  self.skipWaiting();
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var u = new URL(e.request.url);
  if (e.request.mode !== 'navigate' || u.origin !== location.origin || u.pathname !== '/app.html') return;
  e.respondWith(fetch(e.request).catch(function () { return caches.match('/app.html'); }));
});
