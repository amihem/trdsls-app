// Minimal app-shell service worker — the app's actual data lives in Supabase (not in this cache),
// so this only helps the app load a little faster / work briefly offline. It is NETWORK-FIRST:
// it always tries to fetch the latest version first, and only falls back to the cached copy if
// the network is unreachable. This avoids ever serving a stale/old version of the app after a
// new deploy (the earlier cache-first version caused exactly that — needing to clear cache to see
// updates). Bump CACHE_NAME whenever you want to force old cached entries to be dropped.
const CACHE_NAME = "amihem-sales-v2";
const APP_SHELL = ["/", "/index.html", "/manifest.json"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  // Never intercept Supabase API/auth calls — they must always go straight to the network.
  if (e.request.url.includes("supabase.co")) return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        const resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(e.request, resClone));
        return res;
      })
      .catch(() => caches.match(e.request).then((cached) => cached || caches.match("/index.html")))
  );
});
