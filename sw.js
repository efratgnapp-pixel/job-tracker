'use strict';

const CACHE = 'job-tracker-v1';

const PRECACHE = [
  '/job-tracker.html',
  '/interviews.html',
  '/rejections.html',
  '/dashboard.html',
  '/archive.html',
  '/login.html',
  '/icon-192.png',
  '/icon-512.png',
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE).then(c => c.addAll(PRECACHE))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const { pathname } = new URL(event.request.url);

  // Never intercept auth or API calls — always go to network
  if (pathname.startsWith('/api/') || pathname.startsWith('/auth/')) return;

  // Stale-while-revalidate: serve cache immediately, update in background
  event.respondWith(
    caches.open(CACHE).then(async cache => {
      const cached = await cache.match(event.request);
      const networkFetch = fetch(event.request).then(res => {
        if (res.ok) cache.put(event.request, res.clone());
        return res;
      }).catch(() => null);

      return cached || await networkFetch || new Response('Offline', { status: 503 });
    })
  );
});
