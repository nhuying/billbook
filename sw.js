// บันทึกบิล — service worker: ให้แอปเปิดได้แม้ออฟไลน์
const SHELL_CACHE = 'billbook-shell-v1';
const CDN_CACHE = 'billbook-cdn-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(SHELL_CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== SHELL_CACHE && k !== CDN_CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === location.origin) {
    // หน้าแอป: ลองเน็ตก่อน (ได้เวอร์ชันใหม่) ถ้าไม่มีเน็ตใช้ที่เก็บไว้
    if (req.mode === 'navigate' || url.pathname.endsWith('/index.html')) {
      e.respondWith(fetch(req).then(res => {
        const copy = res.clone();
        caches.open(SHELL_CACHE).then(c => c.put('./index.html', copy));
        return res;
      }).catch(() => caches.match('./index.html')));
      return;
    }
    e.respondWith(caches.match(req).then(hit => hit || fetch(req)));
    return;
  }
  // ตัวอ่านสลิปและข้อมูลภาษาไทยจาก CDN: เก็บไว้ใช้ตอนออฟไลน์
  if (url.hostname === 'cdn.jsdelivr.net') {
    e.respondWith(caches.open(CDN_CACHE).then(c => c.match(req).then(hit => hit || fetch(req).then(res => {
      if (res.ok || res.type === 'opaque') c.put(req, res.clone());
      return res;
    }))));
  }
});
