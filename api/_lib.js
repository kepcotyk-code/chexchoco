// 서버 함수 공용 도구 (파일 이름이 _로 시작하므로 Vercel에서 따로 주소가 생기지 않음)
// - Supabase REST 호출
// - 웹 푸시 발송 (VAPID 서명 + RFC 8291 aes128gcm 암호화를 Node 기본 crypto로 직접 구현 → 추가 설치 패키지 없음)
import crypto from 'crypto';

/* ---------------- Supabase ---------------- */
export function supabaseEnv() {
  const url = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || process.env.REACT_APP_SUPABASE_URL || process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY || process.env.REACT_APP_SUPABASE_ANON_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
  if (!url || !key) return null;
  return { url: url.replace(/\/$/, ''), key };
}

export async function sb(env, method, path, body, prefer) {
  const headers = { apikey: env.key, Authorization: `Bearer ${env.key}`, 'Content-Type': 'application/json' };
  if (prefer) headers.Prefer = prefer;
  const res = await fetch(`${env.url}/rest/v1/${path}`, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) });
  const text = await res.text();
  if (!res.ok) throw new Error(`Supabase ${method} ${path.split('?')[0]} 실패 (${res.status}): ${text.slice(0, 200)}`);
  return text ? JSON.parse(text) : null;
}

/* ---------------- base64url ---------------- */
export const b64url = (buf) => Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
export const fromB64url = (s) => Buffer.from(String(s).replace(/-/g, '+').replace(/_/g, '/'), 'base64');

/* ---------------- VAPID ---------------- */
export function vapidEnv() {
  const publicKey = process.env.VAPID_PUBLIC_KEY;
  const privateKey = process.env.VAPID_PRIVATE_KEY;
  if (!publicKey || !privateKey) return null;
  return { publicKey, privateKey, subject: process.env.VAPID_SUBJECT || 'https://chexchoco.vercel.app' };
}

export function createVapidJwt(audience, subject, publicKeyB64, privateKeyB64, nowSec = Math.floor(Date.now() / 1000)) {
  const header = b64url(JSON.stringify({ typ: 'JWT', alg: 'ES256' }));
  const claims = b64url(JSON.stringify({ aud: audience, exp: nowSec + 12 * 3600, sub: subject }));
  const unsigned = `${header}.${claims}`;
  const pub = fromB64url(publicKeyB64);
  const key = crypto.createPrivateKey({
    key: { kty: 'EC', crv: 'P-256', x: b64url(pub.subarray(1, 33)), y: b64url(pub.subarray(33, 65)), d: privateKeyB64 },
    format: 'jwk',
  });
  const sig = crypto.sign('sha256', Buffer.from(unsigned), { key, dsaEncoding: 'ieee-p1363' });
  return `${unsigned}.${b64url(sig)}`;
}

/* ---------------- RFC 8291 (aes128gcm) 암호화 ---------------- */
const hmac = (key, data) => crypto.createHmac('sha256', key).update(data).digest();

// opts.asPrivateKey / opts.salt 는 테스트(표준 예시값 검증)용 - 실제 발송 때는 매번 새로 생성
export function encryptPayload(plaintext, uaPublicB64, authSecretB64, opts = {}) {
  const uaPublic = fromB64url(uaPublicB64);
  const authSecret = fromB64url(authSecretB64);
  const ecdh = crypto.createECDH('prime256v1');
  if (opts.asPrivateKey) ecdh.setPrivateKey(opts.asPrivateKey); else ecdh.generateKeys();
  const asPublic = ecdh.getPublicKey();
  const sharedSecret = ecdh.computeSecret(uaPublic);
  const salt = opts.salt || crypto.randomBytes(16);

  const prkKey = hmac(authSecret, sharedSecret);
  const keyInfo = Buffer.concat([Buffer.from('WebPush: info\0'), uaPublic, asPublic, Buffer.from([1])]);
  const ikm = hmac(prkKey, keyInfo);
  const prk = hmac(salt, ikm);
  const cek = hmac(prk, Buffer.concat([Buffer.from('Content-Encoding: aes128gcm\0'), Buffer.from([1])])).subarray(0, 16);
  const nonce = hmac(prk, Buffer.concat([Buffer.from('Content-Encoding: nonce\0'), Buffer.from([1])])).subarray(0, 12);

  const cipher = crypto.createCipheriv('aes-128-gcm', cek, nonce);
  const padded = Buffer.concat([Buffer.from(plaintext), Buffer.from([2])]); // 마지막 레코드 구분값 0x02
  const encrypted = Buffer.concat([cipher.update(padded), cipher.final(), cipher.getAuthTag()]);

  const header = Buffer.alloc(16 + 4 + 1);
  salt.copy(header, 0);
  header.writeUInt32BE(4096, 16);
  header.writeUInt8(asPublic.length, 20);
  return Buffer.concat([header, asPublic, encrypted]);
}

// 결과: { ok, status, gone } - gone=true면 만료된 구독(삭제 대상)
export async function sendWebPush(sub, payload, vapid) {
  const endpointUrl = new URL(sub.endpoint);
  const jwt = createVapidJwt(`${endpointUrl.protocol}//${endpointUrl.host}`, vapid.subject, vapid.publicKey, vapid.privateKey);
  const body = encryptPayload(Buffer.from(JSON.stringify(payload)), sub.p256dh, sub.auth);
  const res = await fetch(sub.endpoint, {
    method: 'POST',
    headers: {
      TTL: '86400',
      Urgency: 'normal',
      'Content-Encoding': 'aes128gcm',
      'Content-Type': 'application/octet-stream',
      Authorization: `vapid t=${jwt}, k=${vapid.publicKey}`,
    },
    body,
  });
  return { ok: res.status >= 200 && res.status < 300, status: res.status, gone: res.status === 404 || res.status === 410 };
}

/* ---------------- 알림 → 푸시 ---------------- */
// link_id가 'tab:dashboard' 형식이면 그 탭으로, 아니면(도서공유 등) 서재 탭으로 열림
export const linkToUrl = (linkId) => (linkId && String(linkId).startsWith('tab:') ? `/#${String(linkId).slice(4)}` : '/#gallery');

// 알림 내용에서 "OOO님이" 형태로 들어있는 이름을 찾아, 푸시 제목에 "책스초코 · OOO님" 처럼 보낸 사람을 표시함
// (자동 알림처럼 특정 회원이 보낸 게 아닌 경우는 이름이 없으니 그냥 "책스초코"로 남음)
export const pushTitleFor = (message) => {
  const m = String(message || '').match(/([^\s"']{1,12})님이/);
  return m ? `책스초코(${m[1]})` : '책스초코';
};

// 최근 30분 안에 생긴 알림 중 아직 푸시를 안 보낸 것만 골라서, 받는 사람의 모든 기기로 발송
export async function flushPendingPushes(env, vapid) {
  const since = new Date(Date.now() - 30 * 60 * 1000).toISOString();
  // 먼저 pushed_at을 채워서 "내가 보낼 몫"을 확보 (동시에 여러 번 호출돼도 같은 알림이 두 번 가지 않게)
  const claimed = await sb(env, 'PATCH',
    `notifications?pushed_at=is.null&created_at=gte.${encodeURIComponent(since)}&select=id,member_id,message,link_id`,
    { pushed_at: new Date().toISOString() }, 'return=representation');
  if (!claimed || claimed.length === 0) return { sent: 0, claimed: 0 };
  const memberIds = [...new Set(claimed.map((n) => n.member_id))];
  const subs = await sb(env, 'GET', `push_subscriptions?member_id=in.(${memberIds.map((id) => `"${id}"`).join(',')})&select=endpoint,member_id,p256dh,auth`);
  // 홈 화면 아이콘 배지 숫자를 "지금 실제로 안 읽은 개수"로 정확히 맞추기 위해, 회원별 안 읽은 알림 수를 서버에서 직접 세어서 같이 보냄
  // (화면을 안 열어둔 상태에서 알림이 와도, 서비스워커가 이 숫자로 바로 배지를 갱신해서 실제 개수와 어긋나지 않게 함)
  // 회원 수만큼 쿼리를 따로 날리지 않고, 대상 회원들의 안 읽은 알림을 한 번에 받아와서 회원별로 세기만 함(쿼리 1번)
  const unreadCounts = {};
  memberIds.forEach((id) => { unreadCounts[id] = 0; }); // 안 읽은 알림이 하나도 없는 회원도 배지가 0으로 명확히 찍히도록 미리 채워둠
  // silent(알림함에 안 보이는 자동 알림)은 홈 화면 아이콘 배지 숫자에서도 제외 - 앱을 열어봐도 안 보이는 숫자가 배지에 잡히면 헷갈리므로
  const unreadRows = await sb(env, 'GET', `notifications?member_id=in.(${memberIds.map((id) => `"${id}"`).join(',')})&read_at=is.null&silent=eq.false&select=member_id`);
  for (const row of (unreadRows || [])) unreadCounts[row.member_id] = (unreadCounts[row.member_id] || 0) + 1;
  let sent = 0;
  const goneEndpoints = [];
  for (const n of claimed) {
    for (const s of (subs || []).filter((x) => x.member_id === n.member_id)) {
      try {
        const r = await sendWebPush(s, { title: pushTitleFor(n.message), body: n.message, url: linkToUrl(n.link_id), tag: n.id, badge_count: unreadCounts[n.member_id] ?? undefined }, vapid);
        if (r.ok) sent += 1;
        if (r.gone) goneEndpoints.push(s.endpoint);
      } catch (e) { /* 한 기기 실패가 다른 기기 발송을 막지 않도록 */ }
    }
  }
  for (const ep of [...new Set(goneEndpoints)]) {
    try { await sb(env, 'DELETE', `push_subscriptions?endpoint=eq.${encodeURIComponent(ep)}`); } catch (e) { /* 무시 */ }
  }
  return { sent, claimed: claimed.length };
}
