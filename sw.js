// ══════════════════════════════════════════════════════════════
// SERVICE WORKER — cachet alleen het "omhulsel" van de app
// (HTML/CSS/JS/iconen), zodat het icoon en de basisopmaak ook
// zonder internet meteen verschijnen. Data komt ALTIJD live uit
// SharePoint — die wordt hier nooit gecachet.
// ══════════════════════════════════════════════════════════════
const CACHE_NAAM = 'rozal-inzien-v3';
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

  // Iconen veranderen zelden — die mogen gewoon uit cache (snel)
  if (url.includes('/icons/')) {
    event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request)));
    return;
  }

  // Alle code (HTML/JS/CSS/manifest): ALTIJD eerst proberen live op te halen,
  // zodat een nieuwe upload direct zichtbaar is. Cache is puur een reserve
  // voor het geval er geen internet is.
  event.respondWith(
    fetch(event.request)
      .then(resp => {
        const kopie = resp.clone();
        caches.open(CACHE_NAAM).then(cache => cache.put(event.request, kopie));
        return resp;
      })
      .catch(() => caches.match(event.request))
  );
});
