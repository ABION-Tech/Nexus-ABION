// NEXUS · ABION — Service Worker
const CACHE = 'nexus-v5';
const ASSETS = ['/'];
const APP_URL = 'https://nexus-abion.vercel.app/'; // TODO: replace with your actual ABION Vercel deployment URL

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Only cache same-origin GET requests; pass through everything else (incl. Apps Script POSTs)
  if (e.request.method !== 'GET' || !e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(
    caches.match(e.request).then(cached => {
      const fresh = fetch(e.request).then(res => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => cached);
      return cached || fresh;
    })
  );
});

// ── Web Push — NEXUS team chat + per-task discussion notifications ──
self.addEventListener('push', event => {
  let data = {};
  try {
    data = event.data ? event.data.json() : {};
  } catch (e) {
    data = { title: 'NEXUS', body: event.data ? event.data.text() : 'New message' };
  }

  const title = data.title || 'NEXUS';
  const options = {
    body: data.body || '',
    icon: 'icon-192.png',
    badge: 'icon-192.png',
    tag: data.tag || 'nexus-chat',
    data: { url: data.url || APP_URL },
    renotify: true,
  };

  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.url) || APP_URL;

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clientList => {
      // If NEXUS is already open, don't reload it to the homepage — hand
      // the target URL to the page itself so it can navigate in-place
      // (open the right task, chat thread, or notification panel) the
      // same way clicking that item inside the app would. Falls back to
      // openWindow only when there's no existing window to talk to.
      for (const client of clientList) {
        if (client.url.indexOf(new URL(url).origin) === 0) {
          client.postMessage({ type: 'nexus-navigate', url: url });
          if ('focus' in client) return client.focus();
          return;
        }
      }
      if (self.clients.openWindow) return self.clients.openWindow(url);
    })
  );
});
