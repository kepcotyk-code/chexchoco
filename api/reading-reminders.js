// Vercel 서버 함수: /api/reading-reminders
// 1) 독서 시작 알림: 회원이 각자 설정한 시각(기본 12:00)이 지나면 "📖 곧 독서 시간이에요" (켜둔 사람에게만, 모임 있는 날에만)
// 2) 독서 종료 알림: 12:55 이후, 오늘 실제로 체크인한 사람에게만 "⏰ 오늘 독서 종료했어요, 수고하셨어요"
// daily-reminders.js와 달리 하루 한 번이 아니라 여러 번(예: 5분 간격) 호출되는 걸 전제로 함 - 매번 "지금이 그 시각을 지났는지"만
// 확인하고, 알림 id를 고정해서(auto-read-start-회원-날짜 / auto-read-end-회원-날짜) 중복 생성은 항상 무시되므로 몇 번을 불러도 안전함.
// ?dry=1 을 붙이면 보내지 않고 보낼 목록만 돌려줌
import { supabaseEnv, vapidEnv, sb, flushPendingPushes } from './_lib.js';

const ATTENDANCE_DAY_TYPES = ['독서일', '토론회']; // App.jsx의 ATTENDANCE_DAY_TYPES와 동일하게 유지해주세요
const END_TIME = '12:55'; // 종료 알림 고정 시각 (한국시간, 조정 불가 - 필요하면 이 값만 바꾸면 됨)
const pad = (n) => String(n).padStart(2, '0');

export default async function handler(req, res) {
  const dry = req.query?.dry === '1';
  if (!dry && process.env.CRON_SECRET && req.headers.authorization !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ error: 'unauthorized' });
  }
  const env = supabaseEnv();
  if (!env) return res.status(200).json({ skipped: 'Supabase 환경변수가 없어요.' });

  try {
    const kst = new Date(Date.now() + 9 * 3600 * 1000); // 한국 시간 기준
    const today = `${kst.getUTCFullYear()}-${pad(kst.getUTCMonth() + 1)}-${pad(kst.getUTCDate())}`;
    const nowHM = `${pad(kst.getUTCHours())}:${pad(kst.getUTCMinutes())}`;
    const nowIso = new Date().toISOString();

    const days = (await sb(env, 'GET', `calendar_days?date=eq.${today}&select=type`)) || [];
    const meetingToday = days.some((d) => ATTENDANCE_DAY_TYPES.includes(d.type));

    const out = [];
    const add = (id, memberId, message) => out.push({ id, member_id: memberId, message, link_id: 'tab:qr', created_at: nowIso });

    if (meetingToday) {
      // 1) 독서 시작 알림 - 회원별 설정 시각을 지금 시각이 지났으면 (같은 날 한 번만, id 고정으로 자동 중복방지)
      const prefs = (await sb(env, 'GET', 'members_reading_prefs?select=id,notify_reading_start,notify_reading_start_time')) || [];
      for (const p of prefs) {
        if (p.notify_reading_start === false) continue;
        const t = (p.notify_reading_start_time || '12:00').slice(0, 5);
        if (nowHM >= t) add(`auto-read-start-${p.id}-${today}`, p.id, '📖 곧 독서 시간이에요');
      }

      // 2) 독서 종료 알림 - 12:55 이후, 오늘 체크인한 사람에게만
      if (nowHM >= END_TIME) {
        const sessions = (await sb(env, 'GET', `sessions?date=eq.${today}&select=id`)) || [];
        if (sessions.length) {
          const ids = sessions.map((s) => `"${s.id}"`).join(',');
          const checkins = (await sb(env, 'GET', `checkins?session_id=in.(${ids})&check_in_at=not.is.null&select=member_id`)) || [];
          const doneIds = [...new Set(checkins.map((c) => c.member_id))];
          for (const mid of doneIds) add(`auto-read-end-${mid}-${today}`, mid, '⏰ 오늘 독서 종료했어요, 수고하셨어요');
        }
      }
    }

    if (dry) return res.status(200).json({ dry: true, today, nowHM, meetingToday, count: out.length, notifications: out.map((n) => ({ to: n.member_id, message: n.message })) });
    if (out.length) await sb(env, 'POST', 'notifications', out, 'resolution=ignore-duplicates,return=minimal');
    const vapid = vapidEnv();
    const push = vapid ? await flushPendingPushes(env, vapid) : { skipped: 'VAPID 키 없음' };
    return res.status(200).json({ today, nowHM, meetingToday, created: out.length, push });
  } catch (e) {
    return res.status(200).json({ error: String(e.message || e) });
  }
}
