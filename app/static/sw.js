// static/sw.js
self.addEventListener('install', (event) => {
    console.log('Service Worker instalado.');
});

self.addEventListener('fetch', (event) => {
    // O Chrome exige que o evento responda com o fetch real 
    // para validar a instalação do PWA.
    event.respondWith(fetch(event.request));
});