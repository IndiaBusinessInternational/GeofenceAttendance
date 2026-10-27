/* IBI Staff Attendance — Service Worker
   Bump CACHE on every release (keep in sync with the on-screen version badge). */
const CACHE = "ibi-attendance-v2.10";
const CORE = [
  "./",
  "./index.html",
  "./manifest.json",
  "./logo.png",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-512-mask.png",
  "./ibi-share.js"
];

self.addEventListener("install", (e) => {
  self.skipWaiting();
  // Fetch each file with cache:"reload" — a plain addAll() goes through the
  // HTTP cache, which can precache the PREVIOUS build's files under the new
  // cache name and ship stale JS behind a bumped version badge.
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => Promise.all(
        CORE.map((u) => fetch(new Request(u, { cache: "reload" }))
          .then((res) => (res && res.ok) ? c.put(u, res) : null)
          .catch(() => null))
      ))
      .catch(() => {})
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;                 // never touch punch POSTs

  const url = new URL(req.url);
  // Always hit the network for the live backend + fonts — never cache these.
  if (/script\.google\.com|googleusercontent\.com|fonts\.googleapis\.com|fonts\.gstatic\.com/.test(url.host)) {
    return;
  }

  // HTML: network-first so updates appear immediately; fall back to cache offline.
  if (req.mode === "navigate" || req.destination === "document") {
    e.respondWith(
      fetch(req)
        .then((res) => { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); return res; })
        .catch(() => caches.match(req).then((r) => r || caches.match("./index.html")))
    );
    return;
  }

  // Static assets: cache-first, then network.
  e.respondWith(
    caches.match(req).then((cached) =>
      cached || fetch(req).then((res) => {
        const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); return res;
      }).catch(() => cached)
    )
  );
});
