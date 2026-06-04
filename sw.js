self.addEventListener('install', (e) => {
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  console.log('Service Worker is active');
});

self.addEventListener('fetch', (e) => {
});
