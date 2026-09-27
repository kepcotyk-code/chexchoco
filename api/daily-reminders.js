// Vercel 서버 함수: /api/daily-reminders  (vercel.json의 cron 설정으로 매일 아침 8시(한국시간)에 자동 실행)
// 1) 도서 반납기한: 하루 전·당일, 기한 초과 1·3·7일째에 대여자에게 (초과 1일째엔 책 주인에게도)
// 2) 벌칙: 수행 예정일 당일 알림 + 금요일에 이번 주 벌칙 대상인데 예정일을 안 정한 사람에게 알림
// 3) 회비: 매달 지정일(기본 25일)에 이번 달 회비 미납자·미납 회식비가 있는 사람에게 알림
// 같은 날 여러 번 실행돼도 알림 id가 고정이라 중복으로 쌓이지 않음. ?dry=1 을 붙이면 보내지 않고 보낼 목록만 돌려줌
import { supabaseEnv, vapidEnv, sb, flushPendingPushes } from './_lib.js';

const DEFAULT_FEATURES = { todaySummary: true, quarterReport: false, praise: true, autoDue: true, autoPenalty: true, autoDues: true, duesDay: 25 };
const pad = (n) => String(n).padStart(2, '0');
const DAY = 86400000;
const ymdUTC = (d) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
const dayDiff = (a, b) => Math.round((Date.parse(`${a}T00:00:00Z`) - Date.parse(`${b}T00:00:00Z`)) / DAY);
const md = (s) => `${parseInt(s.slice(5, 7), 10)}/${parseInt(s.slice(8, 10), 10)}`;
const EXEMPT_EXCUSE_REASONS = ['출장', '휴가'];
const durationMin = (a, b) => (a && b ? Math.round((new Date(b) - new Date(a)) / 60000) : null);
const attendanceEquivalent = (dur) => (dur !== null && dur >= 30 ? 1 : dur !== null && dur >= 15 ? 0.5 : 0);

export default async function handler(req, res) {
  const dry = req.query?.dry === '1';
  if (!dry && process.env.CRON_SECRET && req.headers.authorization !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ error: 'unauthorized' });
  }
  const env = supabaseEnv();
  if (!env) return res.status(200).json({ skipped: 'Supabase 환경변수가 없어요.' });

  try {
    const settings = await sb(env, 'GET', 'settings?key=eq.features&select=value');
    let features = { ...DEFAULT_FEATURES };
    try { if (settings?.[0]?.value) features = { ...features, ...JSON.parse(settings[0].value) }; } catch (e) { /* 기본값 사용 */ }

    const kst = new Date(Date.now() + 9 * 3600 * 1000); // 한국 시간 기준 날짜 계산
    const today = ymdUTC(kst);
    const tomorrow = ymdUTC(new Date(kst.getTime() + DAY));
    const weekday = kst.getUTCDay(); // 0=일 ... 5=금
    const nowIso = new Date().toISOString();
    const members = (await sb(env, 'GET', 'members_public?select=id,name')) || [];
    const nameOf = (id) => members.find((m) => m.id === id)?.name || '';
    const out = [];
    const add = (id, memberId, message, linkId) => { if (memberId) out.push({ id, member_id: memberId, message, link_id: linkId || null, created_at: nowIso }); };

    // 1) 도서 반납기한
    if (features.autoDue) {
      const shares = (await sb(env, 'GET', 'book_shares?status=eq.matched&due_date=not.is.null&select=id,kind,posted_by,matched_by,book_title,due_date')) || [];
      for (const s of shares) {
        const borrower = s.kind === 'offer' ? s.matched_by : s.posted_by;
        const owner = s.kind === 'offer' ? s.posted_by : s.matched_by;
        const d = dayDiff(today, s.due_date); // 양수 = 기한 지남
        if (s.due_date === tomorrow) add(`auto-due1-${s.id}-${s.due_date}`, borrower, `📚 『${s.book_title}』 반납기한이 내일(${md(s.due_date)})이에요.`, s.id);
        if (d === 0) add(`auto-due0-${s.id}-${s.due_date}`, borrower, `📚 오늘은 『${s.book_title}』 반납일이에요.`, s.id);
        if ([1, 3, 7].includes(d)) add(`auto-overdue-${s.id}-${s.due_date}-${d}`, borrower, `⏰ 『${s.book_title}』 반납기한이 ${d}일 지났어요. 책 주인과 반납 일정을 맞춰주세요.`, s.id);
        if (d === 1) add(`auto-overdue-owner-${s.id}-${s.due_date}`, owner, `⏰ 빌려준 『${s.book_title}』 반납기한이 지났어요.${nameOf(borrower) ? ` (대여: ${nameOf(borrower)})` : ''}`, s.id);
      }
    }

    // 2) 벌칙
    if (features.autoPenalty) {
      const pcs = (await sb(env, 'GET', 'penalty_completions?select=id,session_id,member_id,performed_date,confirmed')) || [];
      for (const p of pcs) {
        if (p.performed_date === today && !p.confirmed) add(`auto-pen-${p.id}-${today}`, p.member_id, "🎯 오늘은 벌칙 수행 예정일이에요. 수행 후 현황 탭에서 '완료'를 눌러주세요.", 'tab:dashboard');
      }
      if (weekday === 5) {
        // 이번 주 월~목이 모두 독서일인 "정상 주"에서 출석 환산 1일 미만인 사람 = 벌칙 대상 (앱의 벌칙 계산과 같은 규칙)
        const monday = ymdUTC(new Date(Date.parse(`${today}T00:00:00Z`) - 4 * DAY));
        const dates = [0, 1, 2, 3].map((i) => ymdUTC(new Date(Date.parse(`${monday}T00:00:00Z`) + i * DAY)));
        const cal = (await sb(env, 'GET', `calendar_days?date=in.(${dates.join(',')})&type=eq.${encodeURIComponent('독서일')}&select=date`)) || [];
        if (new Set(cal.map((c) => c.date)).size === 4) {
          const sessions = (await sb(env, 'GET', `sessions?date=in.(${dates.join(',')})&select=id,date`)) || [];
          const checkins = sessions.length ? ((await sb(env, 'GET', `checkins?session_id=in.(${sessions.map((s) => `"${s.id}"`).join(',')})&select=session_id,member_id,check_in_at,check_out_at`)) || []) : [];
          const excuses = (await sb(env, 'GET', `absence_excuses?date=in.(${dates.join(',')})&select=date,member_id,reason`)) || [];
          for (const m of members) {
            const relevant = dates.filter((ds) => !excuses.some((e) => e.date === ds && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason)));
            if (relevant.length === 0) continue;
            const total = relevant.reduce((sum, ds) => {
              const s = sessions.find((x) => x.date === ds);
              const c = s ? checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id) : null;
              return sum + attendanceEquivalent(c ? durationMin(c.check_in_at, c.check_out_at) : null);
            }, 0);
            if (total >= 1) continue;
            const done = pcs.find((p) => p.session_id === monday && p.member_id === m.id);
            if (done?.performed_date || done?.confirmed) continue;
            add(`auto-penset-${monday}-${m.id}`, m.id, '🎯 이번 주 벌칙 대상이에요. 오늘까지 현황 탭에서 수행 예정일을 정해주세요.', 'tab:dashboard');
          }
        }
      }
    }

    // 3) 회비·회식비 (이번 달에 한 명이라도 납부 기록이 있을 때만 = 회비 걷기가 시작된 달만)
    if (features.autoDues && parseInt(today.slice(8, 10), 10) === Number(features.duesDay || 25)) {
      const month = today.slice(0, 7);
      const dues = (await sb(env, 'GET', `dues_payments?month=eq.${month}&select=member_id,paid`)) || [];
      const unpaidDinner = (await sb(env, 'GET', 'dinner_collections?paid=eq.false&select=member_id,amount')) || [];
      const started = dues.some((d) => d.paid);
      for (const m of members) {
        if (started && !dues.some((d) => d.member_id === m.id && d.paid)) {
          add(`auto-dues-${month}-${m.id}`, m.id, `💰 ${parseInt(month.slice(5), 10)}월 회비가 아직 납부 전이에요. 확인 부탁드려요!`, 'tab:notice');
        }
        const dinnerSum = unpaidDinner.filter((c) => c.member_id === m.id).reduce((s, c) => s + Number(c.amount || 0), 0);
        if (dinnerSum > 0) add(`auto-dinner-${month}-${m.id}`, m.id, `🍽️ 미납 회식비 ${dinnerSum.toLocaleString('ko-KR')}원이 있어요. 확인 부탁드려요!`, 'tab:notice');
      }
    }

    if (dry) return res.status(200).json({ dry: true, today, count: out.length, notifications: out.map((n) => ({ to: nameOf(n.member_id), message: n.message })) });
    if (out.length) await sb(env, 'POST', 'notifications', out, 'resolution=ignore-duplicates,return=minimal');
    const vapid = vapidEnv();
    const push = vapid ? await flushPendingPushes(env, vapid) : { skipped: 'VAPID 키 없음' };
    return res.status(200).json({ today, created: out.length, push });
  } catch (e) {
    return res.status(200).json({ error: String(e.message || e) });
  }
}
