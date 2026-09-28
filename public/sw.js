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

// 앱 화면(App.jsx)에서 "안 읽은 알림 0개"가 되면 이걸 보내옴 - 아이콘 배지 계산과 별개로, 안드로이드 알림창에 아직 떠 있는 알림들도 같이 닫아줌
// (일부 안드로이드 런처는 setAppBadge 숫자가 아니라 "안 닫힌 시스템 알림 개수"로 아이콘 배지를 표시해서, 알림창을 안 닫으면 앱 안에서 다 읽어도 배지가 안 사라짐)
self.addEventListener('message', (event) => {
  if (event.data?.type === 'clear-notifications') {
    event.waitUntil((async () => {
      const list = await self.registration.getNotifications();
      list.forEach((n) => n.close());
    })());
  }
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
