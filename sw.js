// Uygulama dosyalarını önbelleğe alır, böylece internet olmadan da açılır.
// Kod değiştiğinde VERSION'ı artırın ki telefonlar yeni sürümü alsın.
const VERSION = 'v1';
const CACHE = 'personel-' + VERSION;
const SHELL = [
  './',
  'index.html',
  'style.css',
  'app.js',
  'manifest.webmanifest',
  'icons/icon.svg',
  'icons/icon-192.png',
  'icons/icon-512.png',
];
const CDN = [
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await c.addAll(SHELL);
    // Kütüphaneler yalnızca Excel içe aktarırken gerekir; alınamazsa kurulum yine de tamamlansın.
    await Promise.all(CDN.map((u) => c.add(u).catch(() => {})));
    self.skipWaiting();
  })());
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    self.clients.claim();
  })());
});

// Kendi dosyalarımız: önce internetten dene (güncel sürüm gelsin), olmazsa önbellekten aç.
// CDN kütüphaneleri sürümü sabit olduğu için doğrudan önbellekten verilir.
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  const url = e.request.url;
  const own = url.startsWith(self.location.origin);
  if (!own && !CDN.includes(url)) return;
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    if (!own) return (await c.match(e.request)) || fetchAndStore(c, e.request);
    try {
      return await fetchAndStore(c, e.request);
    } catch {
      return (await c.match(e.request, { ignoreSearch: true })) || (await c.match('index.html'));
    }
  })());
});

async function fetchAndStore(cache, req) {
  const res = await fetch(req);
  if (res.ok) cache.put(req, res.clone());
  return res;
}
