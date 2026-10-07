// Paisa service worker — makes the app installable and lets it open offline.
// - index.html: network-first (you always get the latest version when online), cached copy when offline.
// - icons/manifest: cache-first.
// - Everything else (Google Apps Script sync, fonts, POST requests) is never touched.
const CACHE = 'paisa-v7';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './badge-96.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;

  if (req.mode === 'navigate' || url.pathname.endsWith('/index.html') || url.pathname.endsWith('/')) {
    e.respondWith(fetch(req, { cache: 'no-store' })
      .then(res => { if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put('./index.html', copy)); } return res; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req)));
});

// Push (Firebase Cloud Messaging, data-only messages). Shows a notification, or tells the open page to show a toast.
self.addEventListener('push', e => {
  let p = {}; try { p = e.data ? e.data.json() : {}; } catch (x) {}
  const d = p.data || p.notification || p;
  const title = d.title || 'Paisa', body = d.body || '';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => {
    const vis = cs.find(c => c.visibilityState === 'visible');
    if (vis) { cs.forEach(c => c.postMessage({ paisaPush: { title, body, tag: d.tag || '' } })); return; }
    return self.registration.showNotification(title, { body, icon: './icon-192.png', badge: './badge-96.png', tag: d.tag || 'paisa-txn', renotify: true, data: { url: d.url || './' } });
  }));
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = (e.notification.data && e.notification.data.url) || './';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => {
    const c = cs[0];
    if (c) { c.postMessage({ paisaOpen: url }); return c.focus(); }
    return self.clients.openWindow(url);
  }));
});
