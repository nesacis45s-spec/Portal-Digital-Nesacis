/* Service Worker - Jurnal Harian Guru Digital
   Naikkan CACHE_VERSION jika Anda menambah/mengganti file pendukung (ikon, dll). */
const CACHE_VERSION = "v1";
const CACHE = "jhgd-" + CACHE_VERSION;
const SHELL = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png"];
const CDN = [
  "https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js",
  "https://cdn.jsdelivr.net/npm/sweetalert2@11",
  "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"
];
const RUNTIME_HOSTS = ["cdn.jsdelivr.net", "fonts.googleapis.com", "fonts.gstatic.com"];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await Promise.all(SHELL.map((u) => cache.add(u).catch(() => {})));
    // Pustaka pihak ketiga: simpan agar aplikasi bisa dibuka tanpa sinyal
    await Promise.all(CDN.map(async (u) => {
      try { const r = await fetch(u, { mode: "no-cors" }); await cache.put(u, r); } catch (e) {}
    }));
    self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k.startsWith("jhgd-") && k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

function cacheable(res) { return res && (res.status === 200 || res.type === "opaque"); }

// Halaman utama: ambil dari jaringan (agar update cepat), jatuh ke cache jika offline/lambat
function networkFirst(req) {
  return new Promise((resolve) => {
    let done = false;
    const fallback = async () => (await caches.match("./index.html")) || (await caches.match("./"));
    const timer = setTimeout(async () => { const r = await fallback(); if (r && !done) { done = true; resolve(r); } }, 4000);
    fetch(req).then((res) => {
      clearTimeout(timer);
      if (cacheable(res)) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put("./index.html", copy)); }
      if (!done) { done = true; resolve(res); }
    }).catch(async () => {
      clearTimeout(timer);
      const r = await fallback();
      if (!done) { done = true; resolve(r || Response.error()); }
    });
  });
}

// Aset statis & pustaka: pakai cache dulu, perbarui di latar belakang
async function staleWhileRevalidate(req) {
  const cache = await caches.open(CACHE);
  const cached = await cache.match(req);
  const net = fetch(req).then((res) => { if (cacheable(res)) cache.put(req, res.clone()); return res; }).catch(() => null);
  return cached || (await net) || Response.error();
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;                 // POST ke server (login/sinkron) tidak disentuh
  const url = new URL(req.url);
  if (url.hostname.endsWith("google.com") && url.pathname.startsWith("/macros")) return;   // API Apps Script: selalu langsung ke jaringan
  if (url.hostname.endsWith("googleusercontent.com")) return;
  if (req.mode === "navigate") { event.respondWith(networkFirst(req)); return; }
  if (url.origin === self.location.origin || RUNTIME_HOSTS.includes(url.hostname)) {
    event.respondWith(staleWhileRevalidate(req));
  }
});
