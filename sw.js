// ══════════════════════════════════════════════════════════════
// SERVICE WORKER — cachet alleen het "omhulsel" van de app
// (HTML/CSS/JS/iconen), zodat het icoon en de basisopmaak ook
// zonder internet meteen verschijnen. Data komt ALTIJD live uit
// SharePoint — die wordt hier nooit gecachet.
// ══════════════════════════════════════════════════════════════
const CACHE_NAAM = 'rozal-inzien-v2';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './css/style.css',
  './js/config.js',
  './js/auth.js',
  './js/graph.js',
  './js/data.js',
  './js/app.js',
  './js/msal-browser.min.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAAM).then(cache => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(namen => Promise.all(namen.filter(n => n !== CACHE_NAAM).map(n => caches.delete(n))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const url = event.request.url;
  // Nooit cachen: Microsoft login/graph verkeer — dat moet altijd live zijn
  if (url.includes('login.microsoftonline.com') || url.includes('graph.microsoft.com') || url.includes('sharepoint.com')) {
    return; // laat gewoon door naar het netwerk
  }
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});
