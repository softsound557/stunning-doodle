const SHELL_CACHE = "shell-v1";
const DATA_CACHE = "data-v1";
const SHELL_FILES = [
  "index.html",
  "app.js",
  "config.js",
  "style.css",
  "manifest.json",
  "icons/icon-192.png",
  "icons/icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(SHELL_CACHE).then((cache) => cache.addAll(SHELL_FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  if (url.origin === "https://graph.microsoft.com") {
    // データ・写真はネットワーク優先。取れなければ前回取得分を返す（オフライン閲覧用）
    event.respondWith(
      fetch(event.request)
        .then((res) => {
          const clone = res.clone();
          caches.open(DATA_CACHE).then((cache) => cache.put(event.request, clone));
          return res;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // アプリ本体（画面・スクリプト）はキャッシュ優先
  event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request)));
});
