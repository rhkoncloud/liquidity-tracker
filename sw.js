/* Service worker — caches ONLY the static app shell so the UI opens offline.
   It deliberately never touches Microsoft Graph or login endpoints, so no financial
   data or tokens are ever written to the cache. */
const CACHE = "liquidity-shell-v1";
const SHELL = ["./", "index.html", "app.js", "config.js", "manifest.webmanifest", "icon-192.png", "icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  // Never cache or intercept data/auth traffic — always go straight to the network.
  if(url.hostname === "graph.microsoft.com" || url.hostname.endsWith("login.microsoftonline.com") || url.hostname.endsWith("login.live.com")){
    return; // default browser handling, no SW involvement
  }
  if(e.request.method !== "GET") return;
  // App shell: network-first (so config/code edits propagate), fall back to cache offline.
  if(url.origin === self.location.origin){
    e.respondWith(
      fetch(e.request).then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        return res;
      }).catch(() => caches.match(e.request))
    );
  }
});
