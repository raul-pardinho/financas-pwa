const CACHE = 'financas-v1.1.0';
const LOCAL = ['./','index.html','styles.css','app.js','manifest.webmanifest','assets/icon.svg'];
self.addEventListener('install', event => event.waitUntil(caches.open(CACHE).then(c=>c.addAll(LOCAL)).then(()=>self.skipWaiting())));
self.addEventListener('activate', event => event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request).then(resp => {
    const copy = resp.clone();
    caches.open(CACHE).then(c => c.put(event.request, copy)).catch(()=>{});
    return resp;
  }).catch(()=>caches.match('./index.html'))));
});