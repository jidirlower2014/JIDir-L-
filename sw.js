// جماعت اسلامی ضلع دیر پائین — Service Worker
// index.html بدلنے پر VERSION بڑھانا ضروری نہیں؛ صفحہ پہلے نیٹ ورک سے آتا ہے، نیٹ نہ ہو تو محفوظ کاپی کھلتی ہے۔
const VERSION = 'v1';
const CORE_CACHE = 'jid-core-' + VERSION;
const RUNTIME_CACHE = 'jid-runtime-' + VERSION;
const CORE = ['./', './index.html', './manifest.webmanifest',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './icons/favicon.png'];

// Google Sheet / Drive کی درخواستیں کبھی کیش نہ ہوں (ہمیشہ تازہ ڈیٹا)
const NEVER_CACHE = ['script.google.com', 'script.googleusercontent.com', 'drive.google.com', 'docs.google.com'];
// فونٹ اور PDF لائبریریاں آف لائن کے لیے کیش ہوں
const RUNTIME_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com', 'cdnjs.cloudflare.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CORE_CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CORE_CACHE && k !== RUNTIME_CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function timeout(ms) { return new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), ms)); }
function cacheable(res) { return res && (res.ok || res.type === 'opaque'); }

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (NEVER_CACHE.some(h => url.hostname === h || url.hostname.endsWith('.' + h))) return;

  // صفحہ کھولنا: پہلے نیٹ ورک (تازہ ورژن)، ناکام ہو تو محفوظ کاپی
  if (req.mode === 'navigate') {
    e.respondWith(
      Promise.race([fetch(req), timeout(5000)])
        .then(res => {
          const copy = res.clone();
          caches.open(CORE_CACHE).then(c => c.put('./index.html', copy));
          return res;
        })
        .catch(() => caches.match('./index.html').then(r => r || caches.match('./')))
    );
    return;
  }

  // اپنی فائلیں (آئیکن، manifest): پہلے کیش
  if (url.origin === location.origin) {
    e.respondWith(
      caches.match(req).then(hit => hit || fetch(req).then(res => {
        if (cacheable(res)) { const copy = res.clone(); caches.open(CORE_CACHE).then(c => c.put(req, copy)); }
        return res;
      }))
    );
    return;
  }

  // فونٹ/لائبریری: کیش دکھاؤ، ساتھ ہی پس منظر میں تازہ کرو
  if (RUNTIME_HOSTS.includes(url.hostname)) {
    e.respondWith(
      caches.open(RUNTIME_CACHE).then(cache =>
        cache.match(req).then(hit => {
          const net = fetch(req).then(res => { if (cacheable(res)) cache.put(req, res.clone()); return res; }).catch(() => hit);
          return hit || net;
        })
      )
    );
  }
});
