// Vienkāršs service worker PWA instalēšanai.
// Kešo tikai PWA statiskos failus (manifestu un ikonas); pārējie pieprasījumi
// (Streamlit lapa, WebSocket, Gemini) vienmēr iet uz tīklu, lai dati būtu svaigi.
const CACHE = "ligumu-pwa-v1";
const ASSETS = ["manifest.json", "icon-192.png", "icon-512.png", "icon-maskable-512.png", "apple-touch-icon.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  // Tīkls vispirms, kešs tikai kā rezerve, ja nav savienojuma.
  event.respondWith(fetch(req).catch(() => caches.match(req)));
});
