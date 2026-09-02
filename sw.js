// Osar Sonaa Al-Hayah — registration form service worker
// Purpose: cache the app shell (this page + icon + manifest) so the FORM ITSELF
// still opens with no internet connection. Actual submissions still need a live
// connection to reach the Apps Script backend — see the offline queue logic in
// dys_form.html, which stores a submission locally and resends it automatically
// once the connection comes back.
//
// ⚠️ Bump CACHE_NAME any time you meaningfully change what's cached below —
// it forces the activate handler to delete the old cache. That said, the
// fetch strategy below is now NETWORK-FIRST for the HTML page itself, so in
// practice you should rarely need to: whoever has internet always gets the
// latest dys_form.html straight from GitHub, and the cached copy is only
// ever used as an offline fallback (or if the network request is slow/fails).

const CACHE_NAME = "osh-form-shell-v1";
const APP_SHELL = [
  "./osh_form.html",
  "./manifest.webmanifest",
  "./icon.svg",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;

  // Never intercept calls to the Apps Script backend — those must always hit the network.
  if (req.url.includes("script.google.com")) return;

  if (req.method !== "GET") return;

  // The HTML page itself (navigations, and dys_form.html directly): always
  // try the network FIRST so anyone with internet gets the latest version of
  // the form (latest field config, latest bug fixes, etc.) — never a stale
  // cached copy. Only fall back to the cached copy if the network request
  // fails entirely (actually offline), so the offline-opening behavior is
  // preserved. The fresh response also re-populates the cache for next time.
  const isHtmlRequest = req.mode === "navigate" || req.url.endsWith("osh_form.html") || req.url.endsWith("/");
  if (isHtmlRequest) {
    event.respondWith(
      fetch(req)
        .then((res) => {
          if (res.ok) {
            const resClone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, resClone));
          }
          return res;
        })
        .catch(() => caches.match(req))
    );
    return;
  }

  // Everything else (icon, manifest, fonts, ...) rarely changes — cache-first
  // is fine and keeps repeat loads fast.
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req)
        .then((res) => {
          if (res.ok && req.url.startsWith(self.location.origin)) {
            const resClone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, resClone));
          }
          return res;
        })
        .catch(() => cached);
    })
  );
});
