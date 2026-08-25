/* Daily Pulse service worker — new, isolated from MAC-Walkthrough (which
 * has no service worker to extract; see docs/daily-pulse-source-audit.md
 * §24). Own cache name, own scope (this app's own origin/root). Never
 * touches Microsoft identity or Graph requests — those are always
 * network-only, matching source's live-data-only behavior for Daily Pulse.
 */

const CACHE_NAME = "daily-pulse-cache-v1";

const APP_SHELL = [
  "./",
  "index.html",
  "styles.css",
  "config.js",
  "graph.js",
  "auth.js",
  "data.js",
  "roster.js",
  "demo.js",
  "app.js",
  "manifest.webmanifest",
  "icons/icon-192.png",
  "icons/icon-512.png"
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(names => Promise.all(
        names.filter(n => n !== CACHE_NAME).map(n => caches.delete(n))
      ))
      .then(() => self.clients.claim())
  );
});

function isNetworkOnly(url) {
  return url.hostname.endsWith("graph.microsoft.com") ||
         url.hostname.endsWith("login.microsoftonline.com") ||
         url.hostname.endsWith("microsoftonline.com");
}

self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);

  // Never intercept Microsoft identity/Graph traffic — always hit the
  // network directly, no caching, no offline fallback.
  if (isNetworkOnly(url)) return;

  // Only handle same-origin GETs; let everything else pass through.
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return;

  event.respondWith(
    caches.match(event.request).then(cached => {
      const networkFetch = fetch(event.request)
        .then(resp => {
          if (resp && resp.ok) {
            const copy = resp.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
          }
          return resp;
        })
        .catch(() => cached);
      return cached || networkFetch;
    })
  );
});
