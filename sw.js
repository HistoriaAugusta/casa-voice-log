/* Keeps the app on the phone so it opens with no signal. Bump VERSION on
 * every release: the new worker caches the new files and drops the old. */

const VERSION = "v1";
const CACHE = `casa-voice-log-${VERSION}`;
const SHELL = [
  "./",
  "index.html",
  "app.js",
  "style.css",
  "manifest.webmanifest",
  "vendor/msal-browser-3.27.0.min.js",
  "icons/casa-logo.png",
  "icons/apple-touch-icon.png",
  "icons/icon-192.png",
  "icons/icon-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

// App files only; Microsoft sign-in and Graph calls go straight to the network.
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(
    // network first so updates land; the cache answers when there is no signal
    fetch(e.request)
      .then((res) => {
        if (res.ok) {
          const copy = res.clone();   // clone now, before the page reads the body
          caches.open(CACHE).then((c) => c.put(e.request, copy));
        }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true })
        .then((hit) => hit || caches.match("index.html"))),
  );
});
