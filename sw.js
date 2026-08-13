/* سرویس‌ورکر انبار — اپ را روی گوشی نگه می‌دارد تا بدون اینترنت هم باز شود */
const V = 'anbar-v1';
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // درخواست‌های همگام‌سازی هیچ‌وقت کش نمی‌شوند
  if (url.origin !== location.origin) return;
  if (e.request.method !== 'GET') return;

  // اول شبکه، اگر نبود از کش (تا آپدیت‌های اپ زود دیده شود)
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const copy = res.clone();
        caches.open(V).then(c => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request).then(r => r || caches.match('./index.html')))
  );
});
