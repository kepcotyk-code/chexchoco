// Vercel 서버 함수: /api/send-push
// 앱에서 알림(notifications)이 새로 생기면 이 주소를 한 번 불러줌 → 아직 푸시 안 보낸 최근 알림을 받는 사람 휴대폰으로 발송
// 보낼 내용은 요청에서 받지 않고 DB에 실제로 있는 알림만 보내므로, 외부에서 아무 문구나 보내는 용도로 악용할 수 없음
// 필요한 환경변수: VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, SUPABASE_URL(또는 VITE_SUPABASE_URL), SUPABASE_ANON_KEY(또는 VITE_SUPABASE_ANON_KEY)
import { supabaseEnv, vapidEnv, flushPendingPushes } from './_lib.js';

export default async function handler(req, res) {
  const env = supabaseEnv();
  const vapid = vapidEnv();
  if (!env || !vapid) return res.status(200).json({ skipped: '환경변수(Supabase 또는 VAPID 키)가 아직 설정되지 않았어요.' });
  try {
    const result = await flushPendingPushes(env, vapid);
    return res.status(200).json(result);
  } catch (e) {
    return res.status(200).json({ error: String(e.message || e) });
  }
}
