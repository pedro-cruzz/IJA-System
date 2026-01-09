// static/sw.js
self.addEventListener('install', (event) => {
    console.log('Service Worker instalado.');
});

self.addEventListener('fetch', (event) => {

    event.respondWith(fetch(event.request));
});self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => event.waitUntil(self.clients.claim()));
