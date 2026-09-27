// 책스초코 서비스 워커 - 휴대폰 푸시 알림 수신 담당 (화면 캐시는 하지 않음 → 배포하면 항상 최신 화면)
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));

// 서버(/api/send-push)에서 보낸 푸시를 받아 휴대폰 알림으로 띄움
self.addEventListener('push', (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch (e) { data = { body: event.data ? event.data.text() : '' }; }
  const title = data.title || '책스초코';
  event.waitUntil((async () => {
    await self.registration.showNotification(title, {
      body: data.body || '새 알림이 도착했어요.',
      icon: '/icon-192.png',
      badge: '/badge-96.png',
      tag: data.tag || undefined,
      data: { url: data.url || '/#notice' },
    });
    // 홈 화면 아이콘의 빨간 숫자 배지 - 화면을 안 열어도, 서버가 알려준 "진짜 안 읽은 개수"로 바로 맞춤
    if (typeof data.badge_count === 'number' && 'setAppBadge' in self.registration) {
      try {
        if (data.badge_count > 0) await self.registration.setAppBadge(data.badge_count);
        else await self.registration.clearAppBadge();
      } catch (e) { /* 배지 기능 미지원 기기 - 무시 */ }
    }
  })());
});

// 알림을 누르면: 열려 있는 책스초코 화면이 있으면 그 화면으로, 없으면 새로 열기
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const target = new URL((event.notification.data && event.notification.data.url) || '/#notice', self.location.origin);
  const tab = target.hash.replace('#', '') || 'notice';
  event.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    const existing = all.find((c) => c.url.startsWith(self.location.origin));
    if (existing) {
      await existing.focus();
      existing.postMessage({ type: 'open-tab', tab });
      return;
    }
    await self.clients.openWindow(target.href);
  })());
});
