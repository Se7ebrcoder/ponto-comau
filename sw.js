// Guarda os arquivos do app no aparelho para funcionar sem internet.
// Ao publicar uma nova versão, aumente VERSAO para os aparelhos baixarem os arquivos novos.
const VERSAO = "ponto-v5";
const ARQUIVOS = [
  "./", "index.html", "styles.css", "app.js", "xlsm.js", "manifest.webmanifest",
  "vendor/jszip.min.js", "fonts/manrope-latin-wght-normal.woff2",
  "icons/icon-180.png", "icons/icon-192.png", "icons/icon-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSAO).then((c) => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((chaves) => Promise.all(chaves.filter((k) => k !== VERSAO).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then((r) => r || fetch(e.request)));
});
