// static/sw.js
self.addEventListener('install', (event) => {
    console.log('Service Worker instalado.');
});

self.addEventListener('fetch', (event) => {
    // Necessário para habilitar o PWA no Chrome Mobile
});