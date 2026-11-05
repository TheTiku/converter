'use strict';
/**
 * Service worker de Converter: permite instalar la app y usarla sin conexión.
 *  - Los archivos de la app y sus librerías se guardan al instalar.
 *  - Archivos de la app: primero la red (siempre la última versión) y, sin conexión, la copia guardada.
 *  - Librerías e iconos: copia guardada (no cambian).
 *  - Recursos externos (fuentes, motor de OCR, datos de idiomas): se guardan la primera vez que se usan.
 * Al publicar cambios, sube VERSION para renovar las copias guardadas.
 */
const VERSION = 'converter-v1';
const RUNTIME = 'converter-externos';
const CORE = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./css/styles.css",
  "./js/app.js",
  "./js/charts.js",
  "./js/docx.js",
  "./js/editor.js",
  "./js/fonts.js",
  "./js/helpers.js",
  "./js/office.js",
  "./js/pdfconv.js",
  "./js/pptx.js",
  "./js/textlayout.js",
  "./js/tools-extra.js",
  "./js/tools-more.js",
  "./js/tools.js",
  "./vendor/fontkit.umd.min.js",
  "./vendor/html2canvas.min.js",
  "./vendor/jszip.min.js",
  "./vendor/pdf-lib-cantoo.min.js",
  "./vendor/pdf-lib.min.js",
  "./vendor/pdf.min.js",
  "./vendor/pdf.worker.min.js",
  "./vendor/pptxgen.bundle.js",
  "./vendor/tesseract.min.js",
  "./vendor/xlsx.full.min.js",
  "./icons/apple-touch-icon.png",
  "./icons/favicon-32.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png"
];
// fuente más habitual en documentos de Office (Calibri → Carlito): se descarga en segundo plano
const WARM = ['Regular', 'Bold', 'Italic', 'BoldItalic'].map((v) => `https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/carlito/Carlito-${v}.ttf`);

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const cache = await caches.open(VERSION);
    await cache.addAll(CORE);
    self.skipWaiting();
    // calentamiento sin bloquear la instalación
    caches.open(RUNTIME).then((rc) => Promise.all(WARM.map((u) => rc.match(u).then((hit) => hit || fetch(u).then((r) => r.ok && rc.put(u, r)).catch(() => {})))));
  })());
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== VERSION && k !== RUNTIME) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || req.url.startsWith('blob:') || req.url.startsWith('data:')) return;
  const url = new URL(req.url);
  if (url.origin === self.location.origin) {
    if (/\/(vendor|icons)\//.test(url.pathname)) e.respondWith(cacheFirst(req, VERSION));
    else e.respondWith(networkFirst(req));
  } else if (url.protocol === 'https:') {
    e.respondWith(cacheFirst(req, RUNTIME));
  }
});

async function networkFirst(req) {
  const cache = await caches.open(VERSION);
  try {
    const res = await fetch(req.mode === 'navigate' ? req : new Request(req, { cache: 'no-cache' }));
    if (res.ok) cache.put(req.mode === 'navigate' ? './index.html' : req, res.clone());
    return res;
  } catch (_) {
    const hit = await cache.match(req, { ignoreSearch: true }) || (req.mode === 'navigate' && (await cache.match('./index.html')));
    return hit || Response.error();
  }
}

async function cacheFirst(req, name) {
  const cache = await caches.open(name);
  const hit = await cache.match(req, { ignoreSearch: name === VERSION });
  if (hit) return hit;
  const res = await fetch(req);
  if (res.ok && (res.type === 'basic' || res.type === 'cors')) cache.put(req, res.clone());
  return res;
}
