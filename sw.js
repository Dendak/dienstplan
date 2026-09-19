/* Service Worker: immer zuerst das Netz (aktueller Plan), ohne Empfang die zuletzt geladene Version.
   plan.enc bleibt auch im Cache verschlüsselt. */
const CACHE = 'dienstplan-v1';

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())
));

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const cacheable = url.origin === location.origin || url.hostname === 'cdn.jsdelivr.net' || url.hostname.endsWith('gstatic.com') || url.hostname === 'fonts.googleapis.com';
  if (!cacheable) return;
  e.respondWith(
    fetch(req).then(res => {
      if (res.ok || res.type === 'opaque') {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy));
      }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: url.pathname.endsWith('/') || url.pathname.endsWith('.html') })
      .then(hit => hit || (req.mode === 'navigate' ? caches.match('./', { ignoreSearch: true }) : undefined))
      .then(hit => hit || new Response('Offline', { status: 503 })))
  );
});
