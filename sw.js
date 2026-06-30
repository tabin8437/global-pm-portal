var CACHE_NAME = 'pm-portal-v2';
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

  // 우리 앱(같은 출처)의 정적 파일만 캐싱 대상으로 삼는다.
  // 구글 Apps Script API 같은 외부(cross-origin) 요청은 절대 캐싱하지 않고 브라우저가 그대로 처리하게 둔다.
  // (한 번이라도 실패한 API 응답이 캐시되면 계속 그 에러가 재생되는 문제를 막기 위함)
  if (new URL(e.request.url).origin !== self.location.origin) return;

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
