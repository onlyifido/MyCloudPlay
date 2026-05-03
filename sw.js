// 기본적인 PWA 설치를 위한 빈 서비스 워커
self.addEventListener('install', (e) => {
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  console.log('Service Worker is active');
});

self.addEventListener('fetch', (e) => {
  // 네트워크 요청을 가로채지 않고 그대로 통과시킴
});
