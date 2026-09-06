const CACHE = "radar-report-v4-2-clean-20260906";
const CORE = [
  "./", "./index.html", "./styles-v4.2.css", "./app-v4.2.js",
  "./manifest-v4.2.webmanifest", "./icon-192.png", "./icon-512.png"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);

  // HTML/navigation and versioned local JS/CSS use network-first so a newer deploy cannot be hidden by stale cache.
  const local = url.origin === self.location.origin;
  const networkFirst = req.mode === "navigate" || (local && /\.(?:html|js|css|webmanifest)$/.test(url.pathname));

  if (networkFirst) {
    event.respondWith(
      fetch(req).then(resp => {
        const copy = resp.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        return resp;
      }).catch(() => caches.match(req).then(hit => hit || caches.match("./index.html")))
    );
    return;
  }

  event.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(resp => {
      if (local) {
        const copy = resp.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      }
      return resp;
    }))
  );
});
