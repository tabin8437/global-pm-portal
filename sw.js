var CACHE_NAME = 'pm-portal-v1';
var ASSETS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', function(e) {
  e.waitUntil(caches.open(CACHE_NAME).then(function(cache) { return cache.addAll(ASSETS); }));
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.filter(function(k) { return k !== CACHE_NAME; }).map(function(k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  if (e.request.method !== 'GET') return;

  // 페이지 자체(HTML)는 네트워크를 먼저 시도해서 항상 최신 버전을 받아오고,
  // 오프라인일 때만 캐시로 대체 (그 외 정적 파일은 캐시 우선)
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request).then(function(res) {
        var resClone = res.clone();
        caches.open(CACHE_NAME).then(function(cache) { cache.put(e.request, resClone); });
        return res;
      }).catch(function() { return caches.match(e.request); })
    );
    return;
  }

  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request).then(function(res) {
        var resClone = res.clone();
        caches.open(CACHE_NAME).then(function(cache) { cache.put(e.request, resClone); });
        return res;
      }).catch(function() { return cached; });
    })
  );
});
