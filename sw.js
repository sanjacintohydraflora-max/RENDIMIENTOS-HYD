// Service worker: guarda la app en el celular para abrirla sin señal.
// Cambia VERSION en cada publicación para que los celulares bajen la versión nueva.
const VERSION = 'v11';
const CACHE = `rendimientos-${VERSION}`;
const ARCHIVOS = [
  './', 'index.html', 'styles.css', 'manifest.webmanifest',
  'js/db.js', 'js/calc.js', 'js/sync.js', 'js/dashboard.js', 'js/app.js',
  'vendor/xlsx.full.min.js', 'vendor/qrcode.js', 'data/labores.json', 'data/camas.json',
  'icons/logo.png', 'icons/icon-192.png', 'icons/icon-512.png',
];

self.addEventListener('install', ev => {
  // cache: 'reload' salta el caché HTTP del navegador (GitHub Pages lo guarda 10 minutos).
  ev.waitUntil(caches.open(CACHE)
    .then(c => c.addAll(ARCHIVOS.map(u => new Request(u, { cache: 'reload' }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', ev => {
  ev.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k.startsWith('rendimientos-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

// Mismo origen: responde desde el caché y actualiza en segundo plano.
// Otros orígenes (el servidor de Google) pasan directo a la red.
self.addEventListener('fetch', ev => {
  const url = new URL(ev.request.url);
  if (ev.request.method !== 'GET' || url.origin !== location.origin) return;
  ev.respondWith(
    caches.open(CACHE).then(async cache => {
      const enCache = await cache.match(ev.request, { ignoreSearch: true });
      const red = fetch(ev.request, { cache: 'no-cache' }).then(res => {
        if (res.ok) cache.put(ev.request, res.clone());
        return res;
      }).catch(() => null);
      return enCache || (await red) || new Response('Sin conexión', { status: 503 });
    }),
  );
});
