import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './supabaseClient';
import {
  Crown, Shield, Wallet, User, Plus, Pencil, Trash2, Check, X, Lock, AlertCircle,
  Megaphone, QrCode, BarChart3, Users, Settings2, Download, Upload, ChevronLeft, ChevronRight, ChevronUp, ChevronDown,
  LogIn, LogOut, Cake, PartyPopper, Archive, Paperclip, FileText, Eye, Pin, Gavel, BookOpen,
  Image as ImageIcon, Trophy, Plane,
} from 'lucide-react';

/* ---------- design tokens (dark) ---------- */
const INK = '#F2EEE3';
const PAPER_BG = '#141310';
const CARD_BG = '#1E1C16';
const MUTE = '#A39B87';
const LINE = '#332F24';
const ROW_LINE = '#2B2820';
const NEUTRAL_BG = '#26231A';
const NEUTRAL_TEXT = '#C9C2AE';
const INPUT_BG = '#17150F';
const BTN_BG = '#F2EEE3';
const BTN_TEXT = '#161410';

const ROLES = [
  { key: '회장', label: '회장', icon: Crown, ink: '#EFC94C', paper: '#3A2E10' },
  { key: '간사', label: '간사', icon: Shield, ink: '#7FDCCF', paper: '#12302C' },
  { key: '총무', label: '총무', icon: Wallet, ink: '#F0A87C', paper: '#3A2213' },
  { key: '회원', label: '회원', icon: User, ink: '#C7C1B0', paper: '#26231A' },
];
const MANAGE_ROLES = ['회장', '간사', '총무'];
const roleMeta = (role) => ROLES.find((r) => r.key === role) || ROLES[3];
const roleOrder = (role) => { const i = ROLES.findIndex((r) => r.key === role); return i === -1 ? 99 : i; };

const READING_SEGMENT_COLORS = ['#F5DE8A', '#EFC94C', '#D9A93A', '#C99A2E', '#B98A22', '#A97A18'];
// 펼쳐진 책 스티커 모양 — 참고 이미지의 실제 윤곽선 좌표를 추출해 그대로 반영 (상단 V자 + 하단 물결형 스캘럽)
const BOOK_PATH = 'M0.0,51.2 Q0.0,12.0 4.0,8.5 Q8.0,5.0 12.0,3.5 Q16.0,2.0 20.0,1.2 Q24.0,0.5 27.0,0.2 Q30.0,0.0 35.0,0.5 Q40.0,1.0 45.0,2.0 Q50.0,3.0 54.0,5.0 Q58.0,7.0 60.0,8.2 Q62.0,9.5 64.0,8.2 Q66.0,7.0 70.0,5.0 Q74.0,3.0 79.0,1.8 Q84.0,0.5 87.0,0.2 Q90.0,0.0 93.0,0.2 Q96.0,0.5 100.0,1.2 Q104.0,2.0 108.0,3.0 Q112.0,4.0 116.0,7.2 Q120.0,10.5 120.0,51.0 Q120.0,91.6 117.0,93.2 Q114.0,94.7 111.0,93.8 Q108.0,93.0 104.0,92.0 Q100.0,91.0 97.0,90.5 Q94.0,90.0 89.0,90.0 Q84.0,90.0 81.0,90.5 Q78.0,91.0 74.0,92.5 Q70.0,94.0 66.0,96.0 Q62.0,98.0 60.0,98.0 Q58.0,98.0 54.0,96.0 Q50.0,94.0 45.5,92.5 Q41.0,91.0 38.0,90.5 Q35.0,90.0 30.0,90.0 Q25.0,90.0 21.5,90.5 Q18.0,91.0 14.0,92.2 Q10.0,93.5 7.0,94.1 Q4.0,94.7 2.0,92.6 Q0.0,90.5 0.0,51.2 Z';
// 안쪽 점선 스티치 전용 경로 — 단순 축소가 아니라 각 지점에서 테두리와 "일정한 거리"를 유지하도록 계산해서,
// 가운데 V자 노치처럼 오목한 지점에서도 테두리와 점선이 겹치지 않음
const BOOK_STITCH_PATH = 'M6.6,51.4 Q6.6,15.0 9.0,12.9 Q11.4,10.8 14.6,9.6 Q17.8,8.4 21.3,7.7 Q24.9,7.0 27.4,6.8 Q29.9,6.6 34.5,7.1 Q39.0,7.5 43.4,8.4 Q47.8,9.3 51.3,11.0 Q54.8,12.8 58.4,15.0 Q62.0,17.3 65.6,15.0 Q69.2,12.8 72.8,11.0 Q76.3,9.2 80.7,8.1 Q85.1,7.0 87.5,6.8 Q90.0,6.6 92.6,6.8 Q95.1,7.0 98.9,7.7 Q102.6,8.5 105.8,9.2 Q109.0,10.0 111.2,11.8 Q113.4,13.6 113.4,50.6 Q113.4,87.6 113.3,87.6 Q113.3,87.6 111.5,87.1 Q109.7,86.6 105.5,85.6 Q101.3,84.5 97.9,84.0 Q94.5,83.4 89.0,83.4 Q83.5,83.4 79.9,84.0 Q76.3,84.6 71.8,86.3 Q67.4,87.9 63.9,89.7 Q60.4,91.4 60.0,91.4 Q59.6,91.4 56.0,89.6 Q52.5,87.9 47.6,86.2 Q42.6,84.6 39.1,84.0 Q35.5,83.4 30.0,83.4 Q24.5,83.4 20.5,84.0 Q16.5,84.5 12.5,85.8 Q8.4,87.1 7.3,87.3 Q6.3,87.5 6.4,87.7 Q6.6,87.9 6.6,51.4 Z';
const DAY_TYPES = [
  { key: '독서일', label: '독서일', color: '#7FA8D9', bg: '#1E2A38' },
  { key: '휴무일', label: '휴무일', color: '#E0958C', bg: '#3A2420' },
  { key: '토론회', label: '토론회', color: '#D9C24C', bg: '#322D12' },
  { key: '회식일', label: '회식일', color: '#D98A5C', bg: '#332415' },
];
const dayTypeMeta = (key) => DAY_TYPES.find((d) => d.key === key) || null;
const ATTENDANCE_DAY_TYPES = ['독서일', '토론회']; // 출석일자로 산정되는 유형
const WEEKEND_BG = '#302C22'; // 금·토·일 기본(미지정) 배경 — 평일 미지정보다 살짝 밝은 톤
const WEEKEND_TEXT = '#9A9382';

function Stamp({ role, size = 38, tilt = -5 }) {
  const meta = roleMeta(role);
  const Icon = meta.icon;
  return (
    <div className="flex items-center justify-center rounded-full border-2 shrink-0"
      style={{ width: size, height: size, borderColor: meta.ink, color: meta.ink, background: meta.paper, transform: `rotate(${tilt}deg)` }}>
      <Icon size={size * 0.48} strokeWidth={2.25} />
    </div>
  );
}
function RoleChip({ role }) {
  const meta = roleMeta(role);
  return (
    <span className="inline-flex items-center rounded-full px-2.5 py-1 text-xs font-semibold"
      style={{ background: meta.paper, color: meta.ink, fontFamily: "'IBM Plex Mono', monospace" }}>{meta.label}</span>
  );
}
function RolePicker({ value, onChange }) {
  return (
    <div className="grid grid-cols-4 gap-2">
      {ROLES.map((r) => {
        const Icon = r.icon; const active = value === r.key;
        return (
          <button key={r.key} type="button" onClick={() => onChange(r.key)}
            className="flex flex-col items-center gap-1 rounded-xl border-2 py-2.5"
            style={{ borderColor: active ? r.ink : LINE, background: active ? r.paper : CARD_BG, color: active ? r.ink : MUTE }}>
            <Icon size={18} strokeWidth={2.25} />
            <span className="text-[11px] font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{r.label}</span>
          </button>
        );
      })}
    </div>
  );
}
function Card({ children, className = '', style = {} }) {
  return <div className={`rounded-2xl border p-4 ${className}`} style={{ borderColor: LINE, background: CARD_BG, ...style }}>{children}</div>;
}
function PrimaryBtn({ children, onClick, disabled, icon: Icon }) {
  return (
    <button onClick={onClick} disabled={disabled}
      className="flex items-center justify-center gap-1.5 rounded-xl py-2.5 px-4 text-sm font-semibold disabled:opacity-40"
      style={{ background: BTN_BG, color: BTN_TEXT }}>
      {Icon && <Icon size={15} />} {children}
    </button>
  );
}
function GhostBtn({ children, onClick, icon: Icon, color = NEUTRAL_TEXT, bg = NEUTRAL_BG, disabled }) {
  return (
    <button onClick={onClick} disabled={disabled}
      className="flex items-center justify-center gap-1.5 rounded-xl py-2.5 px-4 text-sm font-semibold disabled:opacity-40"
      style={{ background: bg, color }}>
      {Icon && <Icon size={15} />} {children}
    </button>
  );
}
const inputStyle = { borderColor: LINE, background: INPUT_BG, color: INK };

/* ---------- helpers ---------- */
const pad = (n) => String(n).padStart(2, '0');
const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const monthStr = (d = new Date()) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
const fmtTime = (iso) => { if (!iso) return '—'; const d = new Date(iso); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const fmtDate = (s) => { if (!s) return ''; const [y, m, d] = s.slice(0, 10).split('-'); return `${y}.${m}.${d}`; };
const durationMin = (inIso, outIso) => { if (!inIso || !outIso) return null; return Math.round((new Date(outIso) - new Date(inIso)) / 60000); };
// 30분 이상: 1일 인정, 15분 이상 30분 미만: 0.5일 인정, 그 외: 0
const attendanceEquivalent = (dur) => (dur !== null && dur >= 30 ? 1 : dur !== null && dur >= 15 ? 0.5 : 0);
const fmtHM = (totalMin) => { const h = Math.floor(totalMin / 60); const m = totalMin % 60; if (h === 0) return `${m}분`; if (m === 0) return `${h}시간`; return `${h}시간 ${m}분`; };
const mdOf = (birthday) => birthday ? birthday.slice(5, 10) : null;
const fmtMD = (md) => { const [m, d] = md.split('-'); return `${parseInt(m, 10)}월 ${parseInt(d, 10)}일`; };
const maskName = (name) => { if (!name) return name; const chars = [...name]; return chars[0] + 'O'.repeat(Math.max(chars.length - 1, 0)); };
const dispName = (name, loggedIn) => (loggedIn ? name : maskName(name));
const uid = (prefix) => `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
// 오늘 모임장소 위치 확인 (전남 나주시 전력로 55 — 한전 본사 기준)
const MEETING_LAT = 35.0266818;
const MEETING_LNG = 126.7853155;
const MEETING_RADIUS_M = 200;
const MEETING_LABEL = '한전 나주 본사 도서관';
const distanceMeters = (lat1, lng1, lat2, lng2) => {
  const R = 6371000;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
};
// 위치 확인: { ok: true=범위 내 / false=범위 밖 / null=확인 불가, distance: 미터(m) } — 어떤 경우든 체크인 자체는 막지 않음
const getLocationStatus = () => new Promise((resolve) => {
  if (!navigator.geolocation) { resolve({ ok: null, distance: null }); return; }
  navigator.geolocation.getCurrentPosition(
    (pos) => {
      const d = Math.round(distanceMeters(pos.coords.latitude, pos.coords.longitude, MEETING_LAT, MEETING_LNG));
      resolve({ ok: d <= MEETING_RADIUS_M, distance: d });
    },
    () => resolve({ ok: null, distance: null }),
    { enableHighAccuracy: true, timeout: 8000, maximumAge: 0 }
  );
});

const buildMonthGrid = (year, month) => {
  const first = new Date(year, month, 1);
  const startWeekday = first.getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(`${year}-${pad(month + 1)}-${pad(d)}`);
  while (cells.length % 7 !== 0) cells.push(null);
  return cells;
};

// 벌칙 규정: 월~목 4일이 모두 독서일인 "정상 주"에만, 그 4일 중 결석한 날짜가 벌칙 대상이 됨.
// 공휴일 등으로 월~목 중 하루라도 독서일이 아니면 그 주 전체는 벌칙 계산에서 제외.
const isMonToThu = (dateStr) => { const day = new Date(`${dateStr}T00:00:00`).getDay(); return day >= 1 && day <= 4; };
const getMonday = (dateStr) => {
  const d = new Date(`${dateStr}T00:00:00`);
  const diff = d.getDay() === 0 ? -6 : 1 - d.getDay();
  d.setDate(d.getDate() + diff);
  return d;
};
const weekQualifiesForPenalty = (dateStr, calendarDays) => {
  const monday = getMonday(dateStr);
  for (let i = 0; i < 4; i++) {
    const d = new Date(monday); d.setDate(monday.getDate() + i);
    const ds = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    if (!calendarDays.some((c) => c.date === ds && c.type === '독서일')) return false;
  }
  return true;
};
const filterPenaltyEligibleSessions = (sessions, calendarDays) =>
  sessions.filter((s) => isMonToThu(s.date) && weekQualifiesForPenalty(s.date, calendarDays));

const weekKeyOf = (dateStr) => {
  const m = getMonday(dateStr);
  return `${m.getFullYear()}-${pad(m.getMonth() + 1)}-${pad(m.getDate())}`;
};
const EXEMPT_EXCUSE_REASONS = ['출장', '휴가']; // 벌칙 계산에서 제외되는 사유 (개인일정은 제외 안 됨 — 결석으로 그대로 집계)

// 주 단위 벌칙 계산: 월~목 4일이 모두 독서일이고 이미 다 지난 "완결된 주"에서,
// 4일간 출석 환산 합계가 1일 미만(= 30분 이상 출석이 하나도 없고, 15~29분 출석도 2회 미만)인 멤버만 그 주의 벌칙 대상이 됨.
// (15~29분 출석은 0.5일로 환산되므로, 그런 날이 2번이면 1일로 합산되어 벌칙에서 제외됨)
const computeWeeklyPenalties = (sessions, checkins, calendarDays, members, absenceExcuses = []) => {
  const eligible = filterPenaltyEligibleSessions(sessions.filter((s) => s.date < todayStr()), calendarDays);
  const byWeek = {};
  eligible.forEach((s) => { const wk = weekKeyOf(s.date); (byWeek[wk] = byWeek[wk] || []).push(s); });
  return Object.entries(byWeek)
    .filter(([, sess]) => sess.length === 4) // 4일이 다 지나서 세션이 다 있는 주만 (진행 중인 주는 아직 판단 보류)
    .map(([wk, sess]) => {
      const sorted = [...sess].sort((a, b) => a.date.localeCompare(b.date));
      const results = members.map((m) => {
        // 출장/휴가 사유가 있는 날만 그 멤버에 한해 판단 대상에서 제외 (개인일정은 제외 안 됨)
        const relevant = sorted.filter((s) => !absenceExcuses.some((e) => e.date === s.date && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason)));
        if (relevant.length === 0) return { member: m, missedAll: false }; // 4일 다 사유 있으면 벌칙 대상 아님
        const totalEquivalent = relevant.reduce((sum, s) => {
          const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id);
          const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
          return sum + attendanceEquivalent(dur);
        }, 0);
        return { member: m, missedAll: totalEquivalent < 1 };
      });
      return { weekKey: wk, sessions: sorted, results };
    })
    .sort((a, b) => b.weekKey.localeCompare(a.weekKey));
};

/* ---------- Supabase data layer ---------- */
const TABLES = ['members', 'notices', 'notice_views', 'sessions', 'checkins', 'penalty_completions', 'calendar_days', 'settings', 'photos', 'absence_excuses', 'meeting_locations', 'dues_payments', 'expenses', 'dinner_collections'];

async function fetchAll() {
  const results = await Promise.all(TABLES.map((t) => supabase.from(t).select('*')));
  const out = {};
  TABLES.forEach((t, i) => { out[t] = results[i].data || []; });
  return out;
}
async function insertRow(table, row) {
  const { error } = await supabase.from(table).insert(row);
  if (error) throw error;
}
async function updateRow(table, matchCol, matchVal, patch) {
  const { error } = await supabase.from(table).update(patch).eq(matchCol, matchVal);
  if (error) throw error;
}
async function upsertRow(table, row, matchCol) {
  const { error } = await supabase.from(table).upsert(row, { onConflict: matchCol });
  if (error) throw error;
}
async function deleteRow(table, matchCol, matchVal) {
  const { error } = await supabase.from(table).delete().eq(matchCol, matchVal);
  if (error) throw error;
}
function publicUrl(bucket, path) {
  const { data } = supabase.storage.from(bucket).getPublicUrl(path);
  return data.publicUrl;
}

/* ============================================================= */
export default function App() {
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState('');
  const [members, setMembers] = useState([]);
  const [notices, setNotices] = useState([]);
  const [noticeViews, setNoticeViews] = useState([]);
  const [sessions, setSessions] = useState([]);
  const [checkins, setCheckins] = useState([]);
  const [penaltyCompletions, setPenaltyCompletions] = useState([]);
  const [calendarDays, setCalendarDays] = useState([]);
  const [settings, setSettings] = useState([]);
  const [photos, setPhotos] = useState([]);
  const [absenceExcuses, setAbsenceExcuses] = useState([]);
  const [meetingLocations, setMeetingLocations] = useState([]);
  const [duesPayments, setDuesPayments] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [dinnerCollections, setDinnerCollections] = useState([]);
  const [currentUserId, setCurrentUserId] = useState(() => localStorage.getItem('chexchoco-current-user') || null);
  const [tab, setTab] = useState('notice');
  const [showLoginModal, setShowLoginModal] = useState(false);
  const [modalSelectedId, setModalSelectedId] = useState('');
  const [modalPinInput, setModalPinInput] = useState('');
  const [modalPinError, setModalPinError] = useState('');
  const [pendingDelete, setPendingDelete] = useState(null); // { message, onConfirm }
  const [deletePinInput, setDeletePinInput] = useState('');
  const [deletePinError, setDeletePinError] = useState('');

  // 전역 토스트 — 저장/삭제 등 액션 성공·실패를 짧게 알려줌
  const [toast, setToast] = useState(null); // { message, kind: 'success' | 'error' }
  const showToast = (message, kind = 'success') => {
    setToast({ message, kind, key: Date.now() });
  };
  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 2600);
    return () => clearTimeout(t);
  }, [toast]);
  useEffect(() => {
    const handler = () => showToast('요청 처리 중 문제가 발생했어요. 다시 시도해 주세요.', 'error');
    window.addEventListener('unhandledrejection', handler);
    return () => window.removeEventListener('unhandledrejection', handler);
  }, []);

  const reload = async () => {
    try {
      const data = await fetchAll();
      setMembers(data.members); setNotices(data.notices); setNoticeViews(data.notice_views);
      setSessions(data.sessions); setCheckins(data.checkins); setPenaltyCompletions(data.penalty_completions);
      setCalendarDays(data.calendar_days); setSettings(data.settings); setPhotos(data.photos); setAbsenceExcuses(data.absence_excuses); setMeetingLocations(data.meeting_locations); setDuesPayments(data.dues_payments); setExpenses(data.expenses); setDinnerCollections(data.dinner_collections);
      setError('');
    } catch (e) { setError('데이터를 불러오지 못했어요. 새로고침해 주세요.'); }
    setLoaded(true);
  };
  useEffect(() => { reload(); }, []);
  // 오늘 접속자수 집계용 — 페이지 로드마다 방문 기록 1건 남김 (site_visits 테이블, 전체 reload 사이클과는 무관하게 별도 처리)
  useEffect(() => { supabase.from('site_visits').insert({ id: uid('visit'), visited_at: new Date().toISOString() }).then(() => {}).catch(() => {}); }, []);

  const penaltyRule = settings.find((s) => s.key === 'penaltyRule')?.value || '';
  const setPenaltyRule = async (v) => { await upsertRow('settings', { key: 'penaltyRule', value: v }, 'key'); reload(); };

  const setIdentity = (id) => {
    if (id) localStorage.setItem('chexchoco-current-user', id); else localStorage.removeItem('chexchoco-current-user');
    setCurrentUserId(id);
  };
  const openLogin = () => { setShowLoginModal(true); setModalSelectedId(''); setModalPinInput(''); setModalPinError(''); };
  const closeLogin = () => { setShowLoginModal(false); setModalSelectedId(''); setModalPinInput(''); setModalPinError(''); };
  const logout = () => { setIdentity(null); };
  const submitLogin = () => {
    const m = members.find((mm) => mm.id === modalSelectedId);
    if (!m) { setModalPinError('사용자를 선택해 주세요.'); return; }
    if (m.pin && modalPinInput !== m.pin) { setModalPinError('PIN이 일치하지 않아요.'); return; }
    setIdentity(m.id);
    closeLogin();
  };

  const currentMember = members.find((m) => m.id === currentUserId) || null;

  // 삭제 시 실수 방지용 재확인(로그인 PIN) — PIN이 설정된 계정이면 PIN 입력, 아니면 한 번 더 확인만
  const requestDelete = (onConfirm, message = '정말 삭제할까요?') => { setPendingDelete({ onConfirm, message }); setDeletePinInput(''); setDeletePinError(''); };
  const cancelDelete = () => { setPendingDelete(null); setDeletePinInput(''); setDeletePinError(''); };
  const confirmDelete = async () => {
    if (currentMember?.pin && deletePinInput !== currentMember.pin) { setDeletePinError('PIN이 일치하지 않아요.'); return; }
    const action = pendingDelete?.onConfirm;
    cancelDelete();
    if (!action) return;
    try { await action(); showToast('삭제했어요.', 'success'); }
    catch (e) { showToast('삭제에 실패했어요. 다시 시도해 주세요.', 'error'); }
  };
  const noManagerExists = !members.some((m) => MANAGE_ROLES.includes(m.role));
  const canManageUsers = noManagerExists || (currentMember ? MANAGE_ROLES.includes(currentMember.role) : false);
  const canManageAttendance = currentMember ? MANAGE_ROLES.includes(currentMember.role) : false;

  const sortedMembers = useMemo(() => [...members].sort((a, b) => {
    const ro = roleOrder(a.role) - roleOrder(b.role);
    return ro !== 0 ? ro : a.name.localeCompare(b.name, 'ko');
  }), [members]);
  const recentPhotos = useMemo(() => [...photos].sort((a, b) => new Date(b.created_at) - new Date(a.created_at)).slice(0, 6), [photos]);

  const TABS = [
    { key: 'notice', label: '공지', icon: Megaphone },
    { key: 'qr', label: '출석', icon: QrCode },
    { key: 'dashboard', label: '대시보드', icon: BarChart3 },
    { key: 'gallery', label: '포토', icon: ImageIcon },
    { key: 'users', label: '멤버', icon: Users },
    ...(canManageUsers ? [{ key: 'treasury', label: '회계', icon: Wallet }] : []),
    ...(canManageAttendance ? [{ key: 'admin', label: '설정', icon: Settings2 }] : []),
  ];

  if (!loaded) {
    return <div className="min-h-screen flex items-center justify-center" style={{ background: PAPER_BG }}>
      <div className="text-sm" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>불러오는 중…</div>
    </div>;
  }

  return (
    <div className="min-h-screen" style={{ background: PAPER_BG, fontFamily: "'Inter', sans-serif" }}>
      <div className="max-w-3xl mx-auto px-4 pt-6 pb-24">
        <div className="mb-6">
          <div className="flex items-center justify-between mb-1.5">
            <span className="text-[11px] tracking-[0.2em] uppercase" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>KEPCO Reading Club</span>
            <button onClick={() => (currentMember ? logout() : openLogin())}
              className="flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold"
              style={{ background: currentMember ? BTN_BG : NEUTRAL_BG, color: currentMember ? BTN_TEXT : NEUTRAL_TEXT }}>
              {currentMember ? <><Stamp role={currentMember.role} size={16} tilt={0} />{currentMember.name}님 · 로그아웃</> : <>로그인</>}
            </button>
          </div>
          <h1 className="text-center text-4xl font-semibold" style={{ fontFamily: "'Fraunces', serif", color: INK }}>책스초코</h1>
        </div>

        {recentPhotos.length > 0 && (
          <div className="mb-4">
            <div className="grid grid-cols-6 gap-1.5">
              {recentPhotos.map((p) => (
                <button key={p.id} onClick={() => setTab('gallery')} className="aspect-square rounded-lg overflow-hidden" style={{ background: NEUTRAL_BG }}>
                  <img src={publicUrl('photos', p.file_path)} className="w-full h-full object-cover" alt="" loading="lazy" />
                </button>
              ))}
            </div>
          </div>
        )}

        {error && (
          <div className="mb-4 flex items-center gap-2 rounded-xl px-3.5 py-2.5 text-sm" style={{ background: '#3A2213', color: '#F0A87C' }}>
            <AlertCircle size={16} className="shrink-0" />{error}
          </div>
        )}

        {showLoginModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.7)' }} onClick={closeLogin}>
            <div onClick={(e) => e.stopPropagation()} className="w-full max-w-sm rounded-2xl border p-5" style={{ background: CARD_BG, borderColor: LINE }}>
              <div className="flex items-center justify-between mb-4">
                <div className="text-sm font-semibold" style={{ color: INK }}>로그인</div>
                <button onClick={closeLogin}><X size={18} style={{ color: MUTE }} /></button>
              </div>
              {sortedMembers.length === 0 ? (
                <p className="text-sm" style={{ color: MUTE }}>등록된 멤버가 없어요. 사용자관리에서 첫 멤버를 등록해 주세요.</p>
              ) : (
                <div className="space-y-3">
                  <div>
                    <div className="text-xs mb-1" style={{ color: MUTE }}>이름 선택</div>
                    <select value={modalSelectedId} onChange={(e) => { setModalSelectedId(e.target.value); setModalPinInput(''); setModalPinError(''); }}
                      className="w-full rounded-xl border px-3 py-2.5 text-sm outline-none" style={inputStyle}>
                      <option value="">— 선택하세요 —</option>
                      {sortedMembers.map((m) => <option key={m.id} value={m.id}>{m.name} ({m.role})</option>)}
                    </select>
                  </div>
                  {modalSelectedId && (() => {
                    const m = sortedMembers.find((mm) => mm.id === modalSelectedId);
                    return m?.pin ? (
                      <div>
                        <div className="text-xs mb-1" style={{ color: MUTE }}>PIN</div>
                        <input type="password" inputMode="numeric" maxLength={4} value={modalPinInput}
                          onChange={(e) => setModalPinInput(e.target.value.replace(/\D/g, ''))}
                          onKeyDown={(e) => e.key === 'Enter' && submitLogin()}
                          className="w-full rounded-xl border px-3 py-2.5 text-sm tracking-[0.3em] outline-none" style={inputStyle} placeholder="••••" autoFocus />
                      </div>
                    ) : <p className="text-xs" style={{ color: MUTE }}>이 사용자는 PIN이 설정되어 있지 않아요. 바로 로그인할 수 있어요.</p>;
                  })()}
                  {modalPinError && <p className="text-xs" style={{ color: '#F0A87C' }}>{modalPinError}</p>}
                  <PrimaryBtn onClick={submitLogin} icon={LogIn}>로그인</PrimaryBtn>
                </div>
              )}
            </div>
          </div>
        )}

        {pendingDelete && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.7)' }} onClick={cancelDelete}>
            <div onClick={(e) => e.stopPropagation()} className="w-full max-w-sm rounded-2xl border p-5" style={{ background: CARD_BG, borderColor: LINE }}>
              <div className="flex items-center justify-between mb-4">
                <div className="text-sm font-semibold flex items-center gap-1.5" style={{ color: INK }}><Trash2 size={16} style={{ color: '#F0A87C' }} /> 삭제 확인</div>
                <button onClick={cancelDelete}><X size={18} style={{ color: MUTE }} /></button>
              </div>
              <p className="text-sm mb-3" style={{ color: NEUTRAL_TEXT }}>{pendingDelete.message}</p>
              {currentMember?.pin ? (
                <div className="mb-3">
                  <div className="text-xs mb-1" style={{ color: MUTE }}>로그인 PIN 확인</div>
                  <input type="password" inputMode="numeric" maxLength={4} value={deletePinInput}
                    onChange={(e) => { setDeletePinInput(e.target.value.replace(/\D/g, '')); setDeletePinError(''); }}
                    onKeyDown={(e) => e.key === 'Enter' && confirmDelete()}
                    className="w-full rounded-xl border px-3 py-2.5 text-sm tracking-[0.3em] outline-none" style={inputStyle} placeholder="••••" autoFocus />
                  {deletePinError && <p className="text-xs mt-1" style={{ color: '#F0A87C' }}>{deletePinError}</p>}
                </div>
              ) : (
                <p className="text-xs mb-3" style={{ color: MUTE }}>PIN이 설정되어 있지 않아요. 아래 버튼으로 한 번 더 확인해 주세요.</p>
              )}
              <div className="flex gap-2">
                <button onClick={cancelDelete} className="flex-1 rounded-xl py-2.5 text-sm font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>취소</button>
                <button onClick={confirmDelete} className="flex-1 rounded-xl py-2.5 text-sm font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}>삭제</button>
              </div>
            </div>
          </div>
        )}

        {toast && (
          <div key={toast.key} className="fixed left-1/2 z-[60] px-4 py-2.5 rounded-xl text-sm font-semibold shadow-lg"
            style={{ bottom: 24, transform: 'translateX(-50%)', background: toast.kind === 'error' ? '#3A2213' : '#12302C', color: toast.kind === 'error' ? '#F0A87C' : '#7FDCCF' }}>
            {toast.message}
          </div>
        )}

        <div className="flex gap-1.5 mb-4 overflow-x-auto pb-1">
          {TABS.map((t) => {
            const Icon = t.icon; const active = tab === t.key;
            return (
              <button key={t.key} onClick={() => setTab(t.key)}
                className="flex items-center gap-1.5 rounded-full px-3.5 py-2 text-sm font-semibold whitespace-nowrap shrink-0"
                style={{ background: active ? BTN_BG : CARD_BG, color: active ? BTN_TEXT : MUTE, border: `1px solid ${active ? BTN_BG : LINE}` }}>
                <Icon size={15} />{t.label}
              </button>
            );
          })}
        </div>

        {tab === 'notice' && <NoticeScreen notices={notices} noticeViews={noticeViews} currentMember={currentMember} canManage={canManageUsers} reload={reload} members={members} requestDelete={requestDelete} />}
        {tab === 'gallery' && <GalleryScreen photos={photos} currentMember={currentMember} canManage={canManageUsers} reload={reload} members={members} sessions={sessions} checkins={checkins} requestDelete={requestDelete} showToast={showToast} />}
        {tab === 'qr' && <QrScreen members={sortedMembers} currentMember={currentMember} sessions={sessions} checkins={checkins} canManage={canManageUsers} canManageAttendance={canManageAttendance} calendarDays={calendarDays} reload={reload} absenceExcuses={absenceExcuses} meetingLocations={meetingLocations} />}
        {tab === 'dashboard' && <DashboardScreen members={sortedMembers} sessions={sessions} checkins={checkins} penaltyRule={penaltyRule} penaltyCompletions={penaltyCompletions} canManage={canManageUsers} calendarDays={calendarDays} reload={reload} absenceExcuses={absenceExcuses} currentMember={currentMember} />}
        {tab === 'users' && <UsersScreen members={members} sortedMembers={sortedMembers} currentUserId={currentUserId} setIdentity={setIdentity} canManage={canManageUsers} notices={notices} sessions={sessions} checkins={checkins} reload={reload} requestDelete={requestDelete} />}
        {tab === 'treasury' && canManageUsers && <TreasuryScreen members={sortedMembers} duesPayments={duesPayments} expenses={expenses} dinnerCollections={dinnerCollections} currentMember={currentMember} reload={reload} requestDelete={requestDelete} showToast={showToast} />}
        {tab === 'admin' && canManageAttendance && <AdminScreen members={sortedMembers} sessions={sessions} checkins={checkins} penaltyRule={penaltyRule} setPenaltyRule={setPenaltyRule} penaltyCompletions={penaltyCompletions} reload={reload} calendarDays={calendarDays} absenceExcuses={absenceExcuses} requestDelete={requestDelete} currentMember={currentMember} />}
      </div>
    </div>
  );
}

/* ---------------- 공지사항 ---------------- */
const MAX_PDF_BYTES = 3 * 1024 * 1024;

function NoticeScreen({ notices, noticeViews, currentMember, canManage, reload, members, requestDelete }) {
  const [showForm, setShowForm] = useState(false);
  const [title, setTitle] = useState('');
  const [content, setContent] = useState('');
  const [editingId, setEditingId] = useState(null);
  const [pinned, setPinned] = useState(false);
  const [attachedFile, setAttachedFile] = useState(null);
  const [fileError, setFileError] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [expandedViewsId, setExpandedViewsId] = useState(null);
  const isLoggedIn = !!currentMember;

  const todayMd = todayStr().slice(5, 10);
  const birthdayFolksToday = members.filter((m) => m.birthday && mdOf(m.birthday) === todayMd);

  const sorted = [...notices].sort((a, b) => {
    if (!!a.pinned !== !!b.pinned) return a.pinned ? -1 : 1;
    const aOrder = a.sort_order ?? new Date(a.created_at).getTime();
    const bOrder = b.sort_order ?? new Date(b.created_at).getTime();
    return bOrder - aOrder;
  });

  const isPdfSignature = (buf) => { const b = new Uint8Array(buf); return b.length >= 4 && b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46; };
  const extFromType = (type) => {
    if (type === 'application/pdf') return 'pdf';
    if (type === 'image/png') return 'png';
    if (type === 'image/webp') return 'webp';
    return 'jpg';
  };
  const handleFilePick = (e) => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file) return;
    setFileError('');
    if (file.type === 'application/pdf') {
      if (file.size > MAX_PDF_BYTES) { setFileError('3MB 이하의 PDF만 첨부할 수 있어요.'); return; }
      const reader = new FileReader();
      reader.onload = (evt) => {
        if (!isPdfSignature(evt.target.result)) { setFileError('PDF 형식이 아니에요.'); return; }
        setAttachedFile(file);
      };
      reader.readAsArrayBuffer(file.slice(0, 8));
    } else if (file.type.startsWith('image/')) {
      if (file.size > MAX_PHOTO_BYTES) { setFileError('6MB 이하 이미지만 첨부할 수 있어요.'); return; }
      setAttachedFile(file);
    } else {
      setFileError('PDF 또는 이미지 파일만 첨부할 수 있어요.');
    }
  };

  const submit = async () => {
    if (!title.trim() || !content.trim() || submitting) return;
    setSubmitting(true); setFileError('');
    try {
      const id = editingId || uid('n');
      let fileMeta = {};
      if (attachedFile) {
        const path = `${id}.${extFromType(attachedFile.type)}`;
        const { error: upErr } = await supabase.storage.from('notice-files').upload(path, attachedFile, { upsert: true, contentType: attachedFile.type });
        if (upErr) throw upErr;
        fileMeta = { has_file: true, file_name: attachedFile.name, file_type: attachedFile.type, file_uploaded_at: new Date().toISOString() };
      }
      if (editingId) {
        await updateRow('notices', 'id', editingId, { title: title.trim(), content: content.trim(), pinned, ...fileMeta });
      } else {
        await insertRow('notices', { id, title: title.trim(), content: content.trim(), author_name: currentMember?.name || '익명', created_at: new Date().toISOString(), pinned, sort_order: Date.now(), ...fileMeta });
      }
      await reload();
      setTitle(''); setContent(''); setShowForm(false); setEditingId(null); setAttachedFile(null); setPinned(false);
    } catch (e) { setFileError('저장에 실패했어요.'); }
    finally { setSubmitting(false); }
  };
  const startEdit = (n) => { setEditingId(n.id); setTitle(n.title); setContent(n.content); setPinned(!!n.pinned); setAttachedFile(null); setFileError(''); setShowForm(true); };
  const remove = async (id) => { await deleteRow('notices', 'id', id); await deleteRow('notice_views', 'notice_id', id); await reload(); };

  const openAttachment = async (n) => {
    if (currentMember && !noticeViews.some((v) => v.notice_id === n.id && v.member_id === currentMember.id)) {
      await insertRow('notice_views', { id: uid('v'), notice_id: n.id, member_id: currentMember.id, member_name: currentMember.name, viewed_at: new Date().toISOString() });
      reload();
    }
    window.open(publicUrl('notice-files', `${n.id}.${extFromType(n.file_type)}`), '_blank');
  };

  const moveNotice = async (id, direction) => {
    const idx = sorted.findIndex((n) => n.id === id);
    const swapIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (idx < 0 || swapIdx < 0 || swapIdx >= sorted.length) return;
    const a = sorted[idx]; const b = sorted[swapIdx];
    if (!!a.pinned !== !!b.pinned) return; // 고정글과 일반글 사이는 순서 이동 안 함
    const aOrder = a.sort_order ?? new Date(a.created_at).getTime();
    const bOrder = b.sort_order ?? new Date(b.created_at).getTime();
    await updateRow('notices', 'id', a.id, { sort_order: bOrder });
    await updateRow('notices', 'id', b.id, { sort_order: aOrder });
    await reload();
  };

  return (
    <div className="space-y-3">
      {birthdayFolksToday.length > 0 && (
        <Card className="text-center">
          <div className="flex items-center justify-center gap-2 mb-1">
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
            <span className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>오늘은 {birthdayFolksToday.map((m) => dispName(m.name, isLoggedIn)).join(', ')}님 생일이에요!</span>
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
          </div>
          <p className="text-sm" style={{ color: MUTE }}>축하 인사 한마디 건네보는 건 어떨까요 🎂</p>
        </Card>
      )}
      {canManage && (
        showForm ? (
          <Card className="space-y-3">
            <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="제목" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <textarea value={content} onChange={(e) => setContent(e.target.value)} placeholder="내용" rows={4} className="w-full rounded-xl border px-3 py-2 text-sm outline-none resize-none" style={inputStyle} />
            <div>
              <label className="inline-flex items-center gap-1.5 rounded-xl py-2 px-3 text-xs font-semibold cursor-pointer" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
                <Paperclip size={13} /> 파일 첨부 (PDF 3MB / 이미지 6MB 이하)
                <input type="file" accept="application/pdf,image/*" onChange={handleFilePick} className="hidden" />
              </label>
              {attachedFile && <div className="text-xs mt-1.5 flex items-center gap-1" style={{ color: MUTE }}><FileText size={12} /> {attachedFile.name}</div>}
              {fileError && <div className="text-xs mt-1.5" style={{ color: '#F0A87C' }}>{fileError}</div>}
            </div>
            <label className="flex items-center gap-2 text-sm cursor-pointer" style={{ color: INK }}>
              <button type="button" onClick={() => setPinned(!pinned)} className="w-5 h-5 rounded-md border-2 flex items-center justify-center shrink-0"
                style={{ borderColor: pinned ? '#EFC94C' : LINE, background: pinned ? '#3A2E10' : 'transparent' }}>
                {pinned && <Pin size={12} style={{ color: '#EFC94C' }} />}
              </button>
              상단 고정
            </label>
            <div className="flex gap-2"><PrimaryBtn onClick={submit} disabled={submitting} icon={Check}>{submitting ? '저장 중…' : editingId ? '수정 저장' : '게시하기'}</PrimaryBtn>
              <GhostBtn onClick={() => { setShowForm(false); setEditingId(null); setTitle(''); setContent(''); setAttachedFile(null); setFileError(''); setPinned(false); }} icon={X}>취소</GhostBtn></div>
          </Card>
        ) : (
          <button onClick={() => setShowForm(true)} className="w-full flex items-center justify-center gap-1.5 rounded-2xl border-2 border-dashed py-3 text-sm font-semibold" style={{ borderColor: LINE, color: MUTE }}>
            <Plus size={16} /> 공지 작성
          </button>
        )
      )}
      {sorted.length === 0 && <Card><p className="text-sm text-center py-4" style={{ color: MUTE }}>등록된 공지가 없어요.</p></Card>}
      {sorted.map((n) => {
        const views = noticeViews.filter((v) => v.notice_id === n.id);
        const expanded = expandedViewsId === n.id;
        const isImage = n.file_type && n.file_type.startsWith('image/');
        return (
          <Card key={n.id} style={n.pinned ? { borderColor: '#EFC94C', borderWidth: 2 } : {}}>
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="flex items-center gap-1.5">
                  {n.pinned && <Pin size={13} style={{ color: '#EFC94C' }} fill="#EFC94C" />}
                  <h3 className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>{n.title}</h3>
                </div>
                <div className="text-xs mt-0.5" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{dispName(n.author_name, isLoggedIn)} · {fmtDate(n.created_at)} {fmtTime(n.created_at)}</div>
              </div>
              {canManage && (
                <div className="flex gap-1 shrink-0">
                  <button onClick={() => moveNotice(n.id, 'up')} className="p-1.5" style={{ color: MUTE }}><ChevronUp size={15} /></button>
                  <button onClick={() => moveNotice(n.id, 'down')} className="p-1.5" style={{ color: MUTE }}><ChevronDown size={15} /></button>
                  <button onClick={() => startEdit(n)} className="p-1.5" style={{ color: MUTE }}><Pencil size={15} /></button>
                  <button onClick={() => requestDelete(() => remove(n.id), '이 공지사항을 삭제할까요?')} className="p-1.5" style={{ color: '#F0A87C' }}><Trash2 size={15} /></button>
                </div>
              )}
            </div>
            <p className="text-sm mt-2 whitespace-pre-wrap" style={{ color: NEUTRAL_TEXT }}>{n.content}</p>
            {n.has_file && (
              <div className="mt-3">
                {isImage ? (
                  <button onClick={() => openAttachment(n)} className="block w-full rounded-xl overflow-hidden" style={{ background: NEUTRAL_BG }}>
                    <img src={publicUrl('notice-files', `${n.id}.${extFromType(n.file_type)}`)} className="w-full max-h-72 object-cover" alt="" loading="lazy" />
                  </button>
                ) : (
                  <button onClick={() => openAttachment(n)} className="inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
                    <Paperclip size={13} /> {n.file_name || '첨부파일'} 다운로드
                  </button>
                )}
                <div className="flex flex-wrap items-center gap-2 mt-2">
                  {n.file_uploaded_at && <span className="text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>업로드 {fmtDate(n.file_uploaded_at)}</span>}
                  <button onClick={() => canManage && setExpandedViewsId(expanded ? null : n.id)} className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}>
                    <Eye size={12} /> {views.length}명 조회{canManage && views.length > 0 ? (expanded ? ' 숨기기' : ' 보기') : ''}
                  </button>
                </div>
                {canManage && expanded && (
                  <div className="flex flex-wrap gap-1.5 mt-2">
                    {views.map((v) => (
                      <span key={v.id} className="text-[11px] rounded-full px-2 py-1" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
                        {v.member_name} · {fmtDate(v.viewed_at)} {fmtTime(v.viewed_at)}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            )}
          </Card>
        );
      })}
    </div>
  );
}

/* ---------------- 포토로그 ---------------- */
const MAX_PHOTO_BYTES = 6 * 1024 * 1024;
function compressImage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('read failed'));
    reader.onload = (e) => {
      const img = new window.Image();
      img.onerror = () => reject(new Error('decode failed'));
      img.onload = () => {
        const maxDim = 1600;
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width > height) { height = Math.round((height * maxDim) / width); width = maxDim; }
          else { width = Math.round((width * maxDim) / height); height = maxDim; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = width; canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
        canvas.toBlob((blob) => { if (!blob) { reject(new Error('blob failed')); return; } resolve(blob); }, 'image/jpeg', 0.78);
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}
function GalleryScreen({ photos, currentMember, canManage, reload, members, sessions, checkins, requestDelete, showToast }) {
  const [error, setError] = useState('');
  const [uploading, setUploading] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const [editingDate, setEditingDate] = useState(false);
  const [dateInput, setDateInput] = useState('');
  const [editingCaption, setEditingCaption] = useState(false);
  const [captionInput, setCaptionInput] = useState('');
  const isLoggedIn = !!currentMember;
  const sorted = [...photos].sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

  const handleFile = async (e) => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file || !currentMember) return;
    if (!file.type.startsWith('image/')) { setError('이미지 파일만 업로드할 수 있어요.'); return; }
    if (file.size > MAX_PHOTO_BYTES) { setError('6MB 이하 사진만 업로드할 수 있어요.'); return; }
    setError(''); setUploading(true);
    try {
      const blob = await compressImage(file);
      const id = uid('g');
      const path = `${id}.jpg`;
      const { error: upErr } = await supabase.storage.from('photos').upload(path, blob, { contentType: 'image/jpeg' });
      if (upErr) throw upErr;
      await insertRow('photos', { id, uploader_id: currentMember.id, uploader_name: currentMember.name, file_path: path, mime_type: 'image/jpeg', created_at: new Date().toISOString() });
      await reload();
    } catch (e) { setError('업로드에 실패했어요.'); }
    finally { setUploading(false); }
  };
  const removePhoto = async (p) => {
    try { await supabase.storage.from('photos').remove([p.file_path]); } catch (e) {}
    await deleteRow('photos', 'id', p.id);
    await reload();
    setViewingId(null);
  };

  const participantsFor = (photo) => {
    const date = photo.created_at.slice(0, 10);
    const session = sessions.find((s) => s.date === date);
    if (!session) return [];
    const ids = checkins.filter((c) => c.session_id === session.id).map((c) => c.member_id);
    return members.filter((m) => ids.includes(m.id));
  };

  // 캡션 / 게시자 / 그날 참석자 이름으로 검색
  const [searchQuery, setSearchQuery] = useState('');
  const q = searchQuery.trim().toLowerCase();
  const filtered = !q ? sorted : sorted.filter((p) => {
    if ((p.caption || '').toLowerCase().includes(q)) return true;
    if ((p.uploader_name || '').toLowerCase().includes(q)) return true;
    return participantsFor(p).some((m) => m.name.toLowerCase().includes(q));
  });

  const navList = q ? filtered : sorted;
  const viewingIdx = navList.findIndex((p) => p.id === viewingId);
  const viewing = viewingIdx >= 0 ? navList[viewingIdx] : null;
  const showPrev = () => { if (viewingIdx > 0) { setViewingId(navList[viewingIdx - 1].id); setEditingDate(false); setEditingCaption(false); } };
  const showNext = () => { if (viewingIdx >= 0 && viewingIdx < navList.length - 1) { setViewingId(navList[viewingIdx + 1].id); setEditingDate(false); setEditingCaption(false); } };
  const saveDate = async () => {
    if (!viewing || !dateInput) return;
    const time = viewing.created_at.slice(11); // 기존 시각(HH:mm:ss.sssZ)은 그대로 유지
    await updateRow('photos', 'id', viewing.id, { created_at: `${dateInput}T${time}` });
    await reload();
    setEditingDate(false);
  };
  const saveCaption = async () => {
    if (!viewing) return;
    await updateRow('photos', 'id', viewing.id, { caption: captionInput.trim() });
    await reload();
    setEditingCaption(false);
  };

  return (
    <div className="space-y-4">
      <div>
        <label className="flex items-center justify-center gap-1.5 rounded-xl py-2.5 px-4 text-sm font-semibold cursor-pointer w-full"
          style={{ background: currentMember ? NEUTRAL_BG : ROW_LINE, color: currentMember ? NEUTRAL_TEXT : MUTE, opacity: uploading ? 0.6 : 1 }}>
          <ImageIcon size={15} /> {uploading ? '업로드 중…' : '사진 추가'}
          <input type="file" accept="image/*" onChange={handleFile} className="hidden" disabled={!currentMember || uploading} />
        </label>
        {!currentMember && <p className="text-xs mt-1.5" style={{ color: MUTE }}>상단에서 본인을 먼저 선택해야 업로드할 수 있어요.</p>}
        {error && <p className="text-xs mt-1.5" style={{ color: '#F0A87C' }}>{error}</p>}
      </div>
      {photos.length > 0 && (
        <input value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="캡션, 게시자, 참석자 이름으로 검색"
          className="w-full rounded-xl border px-3 py-2.5 text-sm outline-none" style={inputStyle} />
      )}
      {sorted.length === 0 ? (
        <Card><p className="text-sm text-center py-6" style={{ color: MUTE }}>아직 올라온 사진이 없어요.</p></Card>
      ) : filtered.length === 0 ? (
        <Card><p className="text-sm text-center py-6" style={{ color: MUTE }}>검색 결과가 없어요.</p></Card>
      ) : (
        <div className="grid grid-cols-3 gap-2">
          {filtered.map((p) => (
            <button key={p.id} onClick={() => { setViewingId(p.id); setEditingDate(false); setEditingCaption(false); }} className="relative aspect-square rounded-lg overflow-hidden" style={{ background: NEUTRAL_BG }}>
              <img src={publicUrl('photos', p.file_path)} className="w-full h-full object-cover" alt="" loading="lazy" />
              {p.caption && <div className="absolute top-1 right-1.5" style={{ color: 'rgba(255,255,255,0.85)', textShadow: '0 1px 2px rgba(0,0,0,0.6)' }}><FileText size={12} /></div>}
              <div className="absolute bottom-1 left-1.5 right-1.5 flex items-center justify-between text-[10px] font-medium" style={{ color: 'rgba(255,255,255,0.75)', textShadow: '0 1px 2px rgba(0,0,0,0.6)', fontFamily: "'IBM Plex Mono', monospace" }}>
                <span>{p.created_at.slice(0, 10)}</span>
                <span className="truncate ml-1">{dispName(p.uploader_name, isLoggedIn)}</span>
              </div>
            </button>
          ))}
        </div>
      )}
      {viewing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.85)' }} onClick={() => setViewingId(null)}>
          <div className="relative max-w-full max-h-full flex flex-col items-center" onClick={(e) => e.stopPropagation()}>
            {viewingIdx > 0 && (
              <button onClick={showPrev} className="absolute left-0 top-1/2 -translate-y-1/2 -translate-x-1 p-2 rounded-full" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><ChevronLeft size={20} /></button>
            )}
            {viewingIdx < sorted.length - 1 && (
              <button onClick={showNext} className="absolute right-0 top-1/2 -translate-y-1/2 translate-x-1 p-2 rounded-full" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><ChevronRight size={20} /></button>
            )}
            <img src={publicUrl('photos', viewing.file_path)} className="max-w-full max-h-[65vh] rounded-xl" alt="" />
            <div className="flex items-center gap-3 mt-3">
              <span className="text-sm" style={{ color: '#FFFFFF' }}>{dispName(viewing.uploader_name, isLoggedIn)} · {fmtDate(viewing.created_at)} {fmtTime(viewing.created_at)}</span>
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => { setEditingCaption(!editingCaption); setCaptionInput(viewing.caption || ''); }} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><FileText size={15} /></button>
              )}
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => { setEditingDate(!editingDate); setDateInput(viewing.created_at.slice(0, 10)); }} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><Pencil size={15} /></button>
              )}
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => requestDelete(() => removePhoto(viewing), '이 사진을 삭제할까요?')} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#F0A87C' }}><Trash2 size={15} /></button>
              )}
              <button onClick={() => setViewingId(null)} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><X size={15} /></button>
            </div>
            {viewing.caption && !editingCaption && (
              <p className="text-sm mt-2 max-w-sm text-center px-4" style={{ color: 'rgba(255,255,255,0.85)' }}>{viewing.caption}</p>
            )}
            {editingCaption && (
              <div className="flex items-center gap-2 mt-2 w-full max-w-sm px-4" onClick={(e) => e.stopPropagation()}>
                <input value={captionInput} onChange={(e) => setCaptionInput(e.target.value)} placeholder="캡션 추가" className="flex-1 rounded-lg border px-2 py-1.5 text-sm outline-none" style={{ background: '#17150F', borderColor: '#332F24', color: '#F2EEE3' }} />
                <button onClick={saveCaption} className="rounded-lg px-3 py-1.5 text-xs font-semibold" style={{ background: '#F2EEE3', color: '#161410' }}>저장</button>
                <button onClick={() => setEditingCaption(false)} className="rounded-lg px-3 py-1.5 text-xs font-semibold" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}>취소</button>
              </div>
            )}
            {editingDate && (
              <div className="flex items-center gap-2 mt-2" onClick={(e) => e.stopPropagation()}>
                <input type="date" value={dateInput} onChange={(e) => setDateInput(e.target.value)}
                  className="rounded-lg border px-2 py-1.5 text-sm outline-none" style={{ background: '#17150F', borderColor: '#332F24', color: '#F2EEE3' }} />
                <button onClick={saveDate} className="rounded-lg px-3 py-1.5 text-xs font-semibold" style={{ background: '#F2EEE3', color: '#161410' }}>저장</button>
                <button onClick={() => setEditingDate(false)} className="rounded-lg px-3 py-1.5 text-xs font-semibold" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}>취소</button>
              </div>
            )}
            {(() => {
              const people = participantsFor(viewing);
              return people.length > 0 ? (
                <div className="flex flex-wrap justify-center gap-1.5 mt-3 max-w-sm">
                  <span className="text-xs mr-1" style={{ color: 'rgba(255,255,255,0.6)' }}>그날 참석:</span>
                  {people.map((m) => (
                    <span key={m.id} className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs" style={{ background: 'rgba(255,255,255,0.12)', color: '#FFFFFF' }}>
                      {dispName(m.name, isLoggedIn)}
                    </span>
                  ))}
                </div>
              ) : null;
            })()}
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- QR 출결 ---------------- */
function QrScreen({ members, currentMember, sessions, checkins, canManage, canManageAttendance, calendarDays, reload, absenceExcuses, meetingLocations }) {
  const [showLocEdit, setShowLocEdit] = useState(false);
  const [locCustomMode, setLocCustomMode] = useState(false);
  const [locCustomInput, setLocCustomInput] = useState('');
  const todayLocation = meetingLocations.find((l) => l.date === todayStr());
  const setLocation = async (loc) => {
    if (!currentMember || !loc.trim()) return;
    await upsertRow('meeting_locations', { date: todayStr(), location: loc.trim(), updated_by: currentMember.name, updated_at: new Date().toISOString() }, 'date');
    await reload();
    setShowLocEdit(false); setLocCustomMode(false); setLocCustomInput('');
  };
  const today = todayStr();
  const session = sessions.find((s) => s.date === today);
  const [selectedIds, setSelectedIds] = useState([]);
  const [timeInInput, setTimeInInput] = useState('');
  const [timeOutInput, setTimeOutInput] = useState('');
  const myExcuseToday = currentMember ? absenceExcuses.find((e) => e.date === today && e.member_id === currentMember.id) : null;
  const setMyExcuse = async (reason) => {
    if (!currentMember) return;
    await insertRow('absence_excuses', { id: uid('ae'), date: today, member_id: currentMember.id, reason });
    await reload();
  };
  const clearMyExcuse = async () => {
    if (!myExcuseToday) return;
    await deleteRow('absence_excuses', 'id', myExcuseToday.id);
    await reload();
  };

  const startSession = async () => {
    await insertRow('sessions', { id: uid('s'), date: today, created_at: new Date().toISOString() });
    if (!calendarDays.some((d) => d.date === today && d.type === '독서일')) {
      await insertRow('calendar_days', { id: uid('cd'), date: today, type: '독서일' });
    }
    await reload();
  };
  const myCheckin = session ? checkins.find((c) => c.session_id === session.id && c.member_id === currentMember?.id) : null;
  const [locChecking, setLocChecking] = useState(false);
  const [myLocStatus, setMyLocStatus] = useState('checking'); // 'checking' | true | false | null
  useEffect(() => {
    if (!currentMember || !session) return;
    let cancelled = false;
    setMyLocStatus('checking');
    getLocationStatus().then((res) => { if (!cancelled) setMyLocStatus(res); });
    return () => { cancelled = true; };
  }, [currentMember?.id, session?.id]);
  const checkIn = async () => {
    if (!session || !currentMember) return;
    setLocChecking(true);
    const locResult = await getLocationStatus();
    setLocChecking(false);
    await insertRow('checkins', { id: uid('c'), session_id: session.id, member_id: currentMember.id, check_in_at: new Date().toISOString(), check_out_at: null, checkin_loc_ok: locResult.ok });
    await reload();
  };
  const checkOut = async () => {
    if (!myCheckin) return;
    setLocChecking(true);
    const locResult = await getLocationStatus();
    setLocChecking(false);
    await updateRow('checkins', 'id', myCheckin.id, { check_out_at: new Date().toISOString(), checkout_loc_ok: locResult.ok });
    await reload();
  };
  const resetMyCheckin = async () => {
    if (!myCheckin) return;
    if (!window.confirm('오늘 체크인/체크아웃 기록을 초기화할까요?')) return;
    await deleteRow('checkins', 'id', myCheckin.id);
    await reload();
  };
  const todaysCheckins = session ? checkins.filter((c) => c.session_id === session.id) : [];
  const getCheckin = (memberId) => todaysCheckins.find((c) => c.member_id === memberId);
  const checkInMember = async (memberId) => {
    if (!session || getCheckin(memberId)) return;
    await insertRow('checkins', { id: uid('c'), session_id: session.id, member_id: memberId, check_in_at: new Date().toISOString(), check_out_at: null });
    await reload();
  };
  const checkOutMember = async (memberId) => {
    const c = getCheckin(memberId); if (!c || c.check_out_at) return;
    await updateRow('checkins', 'id', c.id, { check_out_at: new Date().toISOString() });
    await reload();
  };

  const toggleSelect = (id) => setSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const selectAll = () => setSelectedIds(members.map((m) => m.id));
  const clearSelect = () => setSelectedIds([]);

  const bulkCheckInSelected = async () => {
    if (!session || selectedIds.length === 0) return;
    const now = new Date().toISOString();
    const additions = selectedIds.filter((id) => !getCheckin(id)).map((id) => ({ id: uid('c'), session_id: session.id, member_id: id, check_in_at: now, check_out_at: null }));
    if (additions.length) await supabase.from('checkins').insert(additions);
    await reload();
  };
  const bulkCheckOutSelected = async () => {
    if (!session || selectedIds.length === 0) return;
    const now = new Date().toISOString();
    const ids = selectedIds.map((id) => getCheckin(id)).filter((c) => c && !c.check_out_at).map((c) => c.id);
    if (ids.length) await supabase.from('checkins').update({ check_out_at: now }).in('id', ids);
    await reload();
  };
  const bulkSetCheckInTime = async () => {
    if (!session || selectedIds.length === 0 || !timeInInput) return;
    const iso = new Date(`${today}T${timeInInput}:00`).toISOString();
    for (const id of selectedIds) {
      const c = getCheckin(id);
      if (c) await updateRow('checkins', 'id', c.id, { check_in_at: iso });
      else await insertRow('checkins', { id: uid('c'), session_id: session.id, member_id: id, check_in_at: iso, check_out_at: null });
    }
    await reload();
  };
  const bulkSetCheckOutTime = async () => {
    if (!session || selectedIds.length === 0 || !timeOutInput) return;
    const iso = new Date(`${today}T${timeOutInput}:00`).toISOString();
    for (const id of selectedIds) {
      const c = getCheckin(id);
      if (c) await updateRow('checkins', 'id', c.id, { check_out_at: iso });
    }
    await reload();
  };

  return (
    <div className="space-y-4">
      <Card style={{ borderColor: '#EFC94C', borderWidth: 1.5, padding: 12 }}>
        <div className="flex items-center gap-2">
          <div className="flex items-center gap-1.5 text-sm font-semibold shrink-0" style={{ color: INK }}>📍 오늘 모임장소</div>
          <div className="flex-1 text-center min-w-0">
            <span className="text-sm font-semibold truncate" style={{ color: todayLocation ? '#EFC94C' : MUTE }}>{todayLocation ? todayLocation.location : '아직 정해지지 않았어요'}</span>
          </div>
          {currentMember && <button onClick={() => setShowLocEdit(!showLocEdit)} className="text-xs underline underline-offset-2 shrink-0" style={{ color: MUTE }}>{todayLocation ? '변경' : '설정'}</button>}
        </div>
        {showLocEdit && currentMember && (
          <div className="mt-3 pt-3 space-y-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
            <div className="flex flex-wrap items-center gap-2">
              <button onClick={() => setLocation('도서관 세미나실')} className="rounded-full px-3 py-1.5 text-xs font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>도서관 세미나실</button>
              <button onClick={() => setLocation('도서관 안쪽 테이블')} className="rounded-full px-3 py-1.5 text-xs font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>도서관 안쪽 테이블</button>
              <button onClick={() => setLocCustomMode(!locCustomMode)} className="rounded-full px-3 py-1.5 text-xs font-semibold" style={{ background: locCustomMode ? '#3A2E10' : NEUTRAL_BG, color: locCustomMode ? '#EFC94C' : NEUTRAL_TEXT }}>수기작성</button>
              {todayLocation && <span className="text-xs" style={{ color: MUTE }}>기존 설정자 : {dispName(todayLocation.updated_by, !!currentMember)}</span>}
            </div>
            {locCustomMode && (
              <div className="flex gap-2">
                <input value={locCustomInput} onChange={(e) => setLocCustomInput(e.target.value)} placeholder="장소 직접 입력" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                <PrimaryBtn onClick={() => setLocation(locCustomInput)} icon={Check}>저장</PrimaryBtn>
              </div>
            )}
          </div>
        )}
      </Card>
      <Card className="text-center" style={{ borderColor: '#7FA8D9', borderWidth: 1.5 }}>
        <div className="text-[11px] uppercase tracking-wider mb-3" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(today)} 오늘의 출결</div>
        {!session ? (
          canManage ? (
            <div className="py-4">
              <p className="text-sm mb-3" style={{ color: MUTE }}>오늘 출결이 아직 시작되지 않았어요.</p>
              <PrimaryBtn onClick={startSession} icon={QrCode}>오늘 출결 시작</PrimaryBtn>
            </div>
          ) : <p className="text-sm py-6" style={{ color: MUTE }}>아직 오늘 출결이 시작되지 않았어요.<br />간사에게 문의해 주세요.</p>
        ) : (
          <div>
            {currentMember ? (
              <div className="mt-4 flex flex-col items-center gap-2">
                {myExcuseToday ? (
                  <div className="flex items-center gap-2">
                    <span className="text-sm rounded-full px-3 py-1.5" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>오늘 사유: {myExcuseToday.reason}</span>
                    <button onClick={clearMyExcuse} className="text-xs underline underline-offset-2" style={{ color: MUTE }}>취소</button>
                  </div>
                ) : !myCheckin ? <PrimaryBtn onClick={checkIn} icon={LogIn} disabled={locChecking}>{locChecking ? '위치 확인 중…' : '체크인'}</PrimaryBtn>
                  : !myCheckin.check_out_at ? (
                    <>
                      <div className="text-sm flex items-center gap-1.5" style={{ color: '#7FDCCF' }}>체크인 {fmtTime(myCheckin.check_in_at)}{myCheckin.checkin_loc_ok === false && <span className="text-[10px] rounded-full px-1.5 py-0.5" style={{ background: '#3A2213', color: '#F0A87C' }}>📍위치 미확인</span>}</div>
                      <div className="flex items-center gap-2">
                        <PrimaryBtn onClick={checkOut} icon={LogOut} disabled={locChecking}>{locChecking ? '위치 확인 중…' : '체크아웃'}</PrimaryBtn>
                        <button onClick={resetMyCheckin} className="text-xs underline underline-offset-2" style={{ color: MUTE }}>초기화</button>
                      </div>
                    </>
                  ) : (
                    <div className="text-sm text-center" style={{ color: MUTE }}>
                      체크인 {fmtTime(myCheckin.check_in_at)}{myCheckin.checkin_loc_ok === false && ' 📍'} → 체크아웃 {fmtTime(myCheckin.check_out_at)}{myCheckin.checkout_loc_ok === false && ' 📍'}
                      {(myCheckin.checkin_loc_ok === false || myCheckin.checkout_loc_ok === false) && <div className="text-[11px] mt-1" style={{ color: '#F0A87C' }}>📍 모임장소 위치가 확인되지 않았어요 (출석 인정에는 영향 없어요)</div>}
                      <button onClick={resetMyCheckin} className="text-xs underline underline-offset-2 mt-2 block mx-auto" style={{ color: MUTE }}>초기화</button>
                    </div>
                  )}
                {!myCheckin && !myExcuseToday && (
                  myLocStatus === 'checking' ? (
                    <span className="text-[11px] rounded-full px-2 py-1 text-center" style={{ background: NEUTRAL_BG, color: MUTE }}>📍 위치 확인 중…</span>
                  ) : myLocStatus.ok === true ? (
                    <span className="text-[11px] rounded-full px-2.5 py-1 text-center" style={{ background: '#12302C', color: '#7FDCCF' }}>📍 {MEETING_LABEL} 기준 약 {myLocStatus.distance}m · {MEETING_RADIUS_M}m 이내라 적정이에요</span>
                  ) : myLocStatus.ok === false ? (
                    <span className="text-[11px] rounded-full px-2.5 py-1 text-center" style={{ background: '#3A2213', color: '#F0A87C' }}>📍 {MEETING_LABEL} 기준 약 {myLocStatus.distance}m · {MEETING_RADIUS_M}m 이내여야 적정이에요</span>
                  ) : (
                    <span className="text-[11px] rounded-full px-2 py-1 text-center" style={{ background: NEUTRAL_BG, color: MUTE }}>📍 위치 확인 불가</span>
                  )
                )}
                {!myExcuseToday && !myCheckin && (
                  <div className="flex items-center gap-1.5 mt-1 flex-wrap justify-center">
                    <span className="text-xs" style={{ color: MUTE }}>오늘 못 오시나요?</span>
                    <button onClick={() => setMyExcuse('출장')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>출장</button>
                    <button onClick={() => setMyExcuse('휴가')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>휴가</button>
                    <button onClick={() => setMyExcuse('개인일정')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>개인일정</button>
                  </div>
                )}
              </div>
            ) : <p className="text-sm mt-3" style={{ color: MUTE }}>상단에서 본인을 먼저 선택해 주세요.</p>}
          </div>
        )}
      </Card>

      {session && canManageAttendance && (
        <Card>
          <div className="flex items-center justify-between mb-2">
            <div className="text-sm font-semibold" style={{ color: INK }}>일괄 출결 관리</div>
            <div className="flex items-center gap-2 text-xs" style={{ color: MUTE }}>
              <span>{selectedIds.length}명 선택</span>
              <button onClick={selectAll} className="underline underline-offset-2">전체 선택</button>
              <button onClick={clearSelect} className="underline underline-offset-2">선택 해제</button>
            </div>
          </div>
          <div className="space-y-1.5 mb-3">
            {members.map((m) => {
              const c = getCheckin(m.id); const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
              const checked = selectedIds.includes(m.id);
              const excuse = absenceExcuses.find((e) => e.date === today && e.member_id === m.id);
              return (
                <div key={m.id} className="flex items-center justify-between text-sm py-1.5" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="flex items-center gap-2 min-w-0">
                    <button onClick={() => toggleSelect(m.id)} className="w-5 h-5 rounded-md border-2 flex items-center justify-center shrink-0"
                      style={{ borderColor: checked ? '#7FDCCF' : LINE, background: checked ? '#12302C' : 'transparent' }}>
                      {checked && <Check size={12} style={{ color: '#7FDCCF' }} />}
                    </button>
                    <Stamp role={m.role} size={26} tilt={0} /><span className="truncate" style={{ color: INK }}>{m.name}</span>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    {c ? <span className="text-xs" style={{ color: dur === null ? MUTE : dur >= 30 ? '#7FDCCF' : '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtTime(c.check_in_at)}–{fmtTime(c.check_out_at)} {dur !== null && `(${dur}분)`}{(c.checkin_loc_ok === false || c.checkout_loc_ok === false) && ' 📍'}</span> : <span className="text-xs" style={{ color: MUTE }}>미체크{excuse ? ` · ${excuse.reason}` : ''}</span>}
                    {!c && <button onClick={() => checkInMember(m.id)} className="p-1.5 rounded-lg" style={{ background: NEUTRAL_BG, color: '#7FDCCF' }}><LogIn size={14} /></button>}
                    {c && !c.check_out_at && <button onClick={() => checkOutMember(m.id)} className="p-1.5 rounded-lg" style={{ background: NEUTRAL_BG, color: '#F0A87C' }}><LogOut size={14} /></button>}
                  </div>
                </div>
              );
            })}
          </div>
          <div className="space-y-2 pt-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
            <div className="flex gap-2">
              <GhostBtn onClick={bulkCheckInSelected} icon={LogIn} bg="#12302C" color="#7FDCCF">선택 체크인</GhostBtn>
              <GhostBtn onClick={bulkCheckOutSelected} icon={LogOut} bg="#3A2213" color="#F0A87C">선택 체크아웃</GhostBtn>
            </div>
            <div className="flex items-center gap-2">
              <input type="time" value={timeInInput} onChange={(e) => setTimeInInput(e.target.value)} className="flex-1 rounded-lg border px-2 py-1.5 text-xs outline-none" style={inputStyle} />
              <GhostBtn onClick={bulkSetCheckInTime} icon={LogIn}>체크인 시간 지정</GhostBtn>
            </div>
            <div className="flex items-center gap-2">
              <input type="time" value={timeOutInput} onChange={(e) => setTimeOutInput(e.target.value)} className="flex-1 rounded-lg border px-2 py-1.5 text-xs outline-none" style={inputStyle} />
              <GhostBtn onClick={bulkSetCheckOutTime} icon={LogOut}>체크아웃 시간 지정</GhostBtn>
            </div>
            <p className="text-[11px]" style={{ color: MUTE }}>선택한 인원에게만 적용돼요. 시간 지정은 체크인 기록이 없어도(체크인은 자동 생성) 적용되고, 체크아웃 시간 지정은 기존 체크인이 있는 인원에만 적용돼요.</p>
          </div>
        </Card>
      )}
      {session && !canManageAttendance && (
        <Card>
          <div className="text-sm font-semibold mb-2" style={{ color: INK }}>오늘 체크인 현황 ({todaysCheckins.length}명)</div>
          <div className="space-y-1.5">
            {todaysCheckins.map((c) => {
              const m = members.find((mm) => mm.id === c.member_id); const dur = durationMin(c.check_in_at, c.check_out_at);
              return (
                <div key={c.id} className="flex items-center justify-between text-sm py-1">
                  <div className="flex items-center gap-2"><Stamp role={m?.role || '회원'} size={24} tilt={0} /><span style={{ color: INK }}>{m ? dispName(m.name, !!currentMember) : '알 수 없음'}</span></div>
                  <span style={{ color: dur === null ? MUTE : dur >= 30 ? '#7FDCCF' : '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtTime(c.check_in_at)}–{fmtTime(c.check_out_at)} {dur !== null && `(${dur}분)`}{(c.checkin_loc_ok === false || c.checkout_loc_ok === false) && ' 📍'}</span>
                </div>
              );
            })}
          </div>
        </Card>
      )}
      {(() => {
        const todaysExcuses = absenceExcuses.filter((e) => e.date === today);
        if (todaysExcuses.length === 0) return null;
        return (
          <Card>
            <div className="text-sm font-semibold mb-2" style={{ color: INK }}>오늘 못 오는 멤버 ({todaysExcuses.length}명)</div>
            <div className="flex flex-wrap gap-1.5">
              {todaysExcuses.map((e) => {
                const m = members.find((mm) => mm.id === e.member_id);
                return (
                  <span key={e.id} className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
                    <Stamp role={m?.role || '회원'} size={16} tilt={0} />{m ? dispName(m.name, !!currentMember) : '알 수 없음'} · {e.reason}
                  </span>
                );
              })}
            </div>
          </Card>
        );
      })()}
    </div>
  );
}

/* ---------------- 대시보드 ---------------- */
function DashboardScreen({ members, sessions, checkins, penaltyRule, penaltyCompletions, canManage, calendarDays, reload, absenceExcuses, currentMember }) {
  const isLoggedIn = !!currentMember;
  const [cursor, setCursor] = useState(new Date());
  const [viewMode, setViewMode] = useState('month');
  const [selectedDate, setSelectedDate] = useState(null);
  const ms = monthStr(cursor);
  const sessionsInMonth = sessions.filter((s) => s.date.startsWith(ms) && s.date <= todayStr());
  const totalDays = sessionsInMonth.length;
  const shift = (delta) => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + delta, 1));

  const sortedSessionsInMonth = [...sessionsInMonth].sort((a, b) => a.date.localeCompare(b.date));
  const sessionWeekKeys = sortedSessionsInMonth.map((s) => weekKeyOf(s.date)); // 도트를 주 단위로 묶어서 표시하기 위함
  const weekChunkRanges = []; // [[startIdx, endIdx), ...] — 한 주(보통 월~목 4일)씩 묶은 구간
  sessionWeekKeys.forEach((wk, i) => {
    if (i === 0 || wk !== sessionWeekKeys[i - 1]) weekChunkRanges.push([i, i + 1]);
    else weekChunkRanges[weekChunkRanges.length - 1][1] = i + 1;
  });
  // 30분 이상: 정상 출석(1일), 15분 이상 30분 미만: 절반 인정(0.5일), 출장/휴가 사유: 별도 표시, 그 외: 결석
  const attendanceStatus = (dur) => (dur !== null && dur >= 30 ? 'full' : dur !== null && dur >= 15 ? 'half' : 'none');
  const rowsUnranked = members.map((m) => {
    const flags = sortedSessionsInMonth.map((s) => {
      const excused = absenceExcuses.some((e) => e.date === s.date && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason));
      if (excused) return 'excused';
      const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id);
      const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
      return attendanceStatus(dur);
    });
    const present = flags.reduce((sum, f) => sum + (f === 'full' ? 1 : f === 'half' ? 0.5 : 0), 0);
    return { ...m, present, flags, rate: totalDays ? Math.round((present / totalDays) * 100) : 0 };
  }).sort((a, b) => b.rate - a.rate);
  let lastRate = null; let lastRank = 0;
  const rows = rowsUnranked.map((r, i) => {
    if (r.rate !== lastRate) { lastRank = i + 1; lastRate = r.rate; }
    return { ...r, rank: lastRank };
  });
  const withRank = (rowsUnranked, key = 'totalMin') => {
    let lastVal = null; let lastRank = 0;
    return rowsUnranked.map((r, i) => {
      if (r[key] !== lastVal) { lastRank = i + 1; lastVal = r[key]; }
      return { ...r, rank: lastRank };
    });
  };
  const totalSessions = sessions.filter((s) => s.date <= todayStr()).length;
  const allTimeRows = withRank(members.map((m) => {
    let present = 0;
    sessions.forEach((s) => { const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id); const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null; const st = attendanceStatus(dur); present += st === 'full' ? 1 : st === 'half' ? 0.5 : 0; });
    return { ...m, present, rate: totalSessions ? Math.round((present / totalSessions) * 100) : 0 };
  }).sort((a, b) => b.rate - a.rate), 'rate');

  const weeklyPenalties = computeWeeklyPenalties(sessions, checkins, calendarDays, members, absenceExcuses);
  const isWeekCompleted = (wk, memberId) => penaltyCompletions.some((p) => p.session_id === wk && p.member_id === memberId);
  const penaltyByMember = {};
  members.forEach((m) => { penaltyByMember[m.id] = { pending: 0 }; });
  weeklyPenalties.forEach((w) => w.results.forEach((r) => {
    if (r.missedAll && !isWeekCompleted(w.weekKey, r.member.id)) penaltyByMember[r.member.id].pending += 1;
  }));

  // 이번 주(진행 중) 지금까지 열린 독서일을 전부 결석한 경우 → 벌칙 예상 경고
  // (하루라도 출석/사유 있으면 그 시점에 바로 해제되어야 하므로, 월·화만 보지 않고 오늘까지의 모든 세션을 확인한다)
  const thisWeekKey = weekKeyOf(todayStr());
  const thisWeekSessionsSoFar = [...sessions]
    .filter((s) => isMonToThu(s.date) && weekKeyOf(s.date) === thisWeekKey && s.date <= todayStr())
    .sort((a, b) => a.date.localeCompare(b.date));
  const thisWeekQualifies = weekQualifiesForPenalty(todayStr(), calendarDays);
  const todayDow = new Date(`${todayStr()}T00:00:00`).getDay(); // 0=일 1=월 2=화 3=수 4=목
  const warningMemberIds = new Set();
  if (thisWeekQualifies && todayDow >= 3 && thisWeekSessionsSoFar.length > 0) {
    members.forEach((m) => {
      const relevant = thisWeekSessionsSoFar.filter((s) => !absenceExcuses.some((e) => e.date === s.date && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason)));
      if (relevant.length === 0) return; // 지금까지의 날짜가 전부 출장/휴가 사유면 경고 대상 아님
      const totalEquivalent = relevant.reduce((sum, s) => {
        const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id);
        const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
        return sum + attendanceEquivalent(dur);
      }, 0);
      if (totalEquivalent < 1) warningMemberIds.add(m.id);
    });
  }

  const isCurrentMonth = ms === monthStr(new Date());

  const monthGrid = buildMonthGrid(cursor.getFullYear(), cursor.getMonth());
  const getDayTypes = (date) => calendarDays.filter((d) => d.date === date).map((d) => d.type);
  const toggleDayType = async (date, type) => {
    const existing = calendarDays.find((d) => d.date === date && d.type === type);
    if (existing) {
      await deleteRow('calendar_days', 'id', existing.id);
      if (ATTENDANCE_DAY_TYPES.includes(type)) {
        const stillCountsForAttendance = calendarDays.some((d) => d.date === date && d.type !== type && ATTENDANCE_DAY_TYPES.includes(d.type));
        if (!stillCountsForAttendance) {
          const s = sessions.find((ss) => ss.date === date);
          if (s && !checkins.some((c) => c.session_id === s.id)) await deleteRow('sessions', 'id', s.id);
        }
      }
    } else {
      await insertRow('calendar_days', { id: uid('cd'), date, type });
      if (ATTENDANCE_DAY_TYPES.includes(type) && !sessions.some((s) => s.date === date)) {
        await insertRow('sessions', { id: uid('s'), date, created_at: new Date().toISOString() });
      }
    }
    await reload();
  };

  return (
    <div className="space-y-4">
      <div className="flex gap-2">
        <button onClick={() => setViewMode('month')} className="flex-1 rounded-xl py-2 text-sm font-semibold" style={{ background: viewMode === 'month' ? BTN_BG : NEUTRAL_BG, color: viewMode === 'month' ? BTN_TEXT : NEUTRAL_TEXT }}>이번 달</button>
        <button onClick={() => setViewMode('all')} className="flex-1 rounded-xl py-2 text-sm font-semibold" style={{ background: viewMode === 'all' ? BTN_BG : NEUTRAL_BG, color: viewMode === 'all' ? BTN_TEXT : NEUTRAL_TEXT }}>전체 기간</button>
      </div>

      {viewMode === 'month' ? (
        <>
          <Card>
            <div className="flex items-center justify-between mb-1">
              <button onClick={() => shift(-1)} className="p-1.5" style={{ color: MUTE }}><ChevronLeft size={18} /></button>
              <div className="flex items-center gap-1.5 font-semibold" style={{ color: INK }}>
                <BookOpen size={15} style={{ color: '#F0A87C' }} />{cursor.getFullYear()}년 {cursor.getMonth() + 1}월
              </div>
              <button onClick={() => shift(1)} className="p-1.5" style={{ color: MUTE }}><ChevronRight size={18} /></button>
            </div>
            <div className="text-xs text-center" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>이번 달 출결 {totalDays}회</div>
          </Card>

          <Card>
            <div className="grid grid-cols-7 gap-1 mb-2">
              {['일', '월', '화', '수', '목', '금', '토'].map((w) => (
                <div key={w} className="text-center text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{w}</div>
              ))}
            </div>
            <div className="grid grid-cols-7 gap-1">
              {monthGrid.map((date, idx) => {
                if (!date) return <div key={idx} />;
                const day = parseInt(date.slice(8, 10), 10);
                const types = getDayTypes(date);
                const hasFeast = types.includes('회식일');
                const fillTypes = types.filter((t) => t !== '회식일'); // 회식일은 테두리로만 표시, 채우기 색에서는 제외
                const metas = fillTypes.map(dayTypeMeta).filter(Boolean);
                const isToday = date === todayStr();
                const dow = new Date(`${date}T00:00:00`).getDay();
                const isWeekendDefault = metas.length === 0 && (dow === 0 || dow === 5 || dow === 6);
                const bgStyle = metas.length === 0
                  ? (hasFeast ? dayTypeMeta('회식일').bg : (isWeekendDefault ? WEEKEND_BG : NEUTRAL_BG))
                  : metas.length === 1 ? metas[0].bg
                  : `linear-gradient(to bottom, ${metas.map((m, i) => `${m.bg} ${(i * 100) / metas.length}%, ${m.bg} ${((i + 1) * 100) / metas.length}%`).join(', ')})`;
                const textColor = metas.length > 0 ? metas[0].color : (hasFeast ? dayTypeMeta('회식일').color : (isWeekendDefault ? WEEKEND_TEXT : MUTE));
                const hasExempt = absenceExcuses.some((e) => e.date === date && EXEMPT_EXCUSE_REASONS.includes(e.reason));
                const hasPersonal = absenceExcuses.some((e) => e.date === date && e.reason === '개인일정');
                const hasDiscussion = types.includes('토론회');
                const birthdayFolks = members.filter((m) => m.birthday && mdOf(m.birthday) === date.slice(5, 10));
                const hasBirthday = birthdayFolks.length > 0;
                let borderStyle = isToday ? `1.5px solid ${INK}` : selectedDate === date ? `1.5px solid ${textColor}` : '1px solid transparent';
                if (hasFeast) borderStyle = '1.5px solid rgba(229, 72, 77, 0.65)';
                return (
                  <button key={date} onClick={() => setSelectedDate(date === selectedDate ? null : date)}
                    className="relative aspect-square rounded-lg flex flex-col items-center justify-center text-xs leading-none"
                    style={{ background: bgStyle, color: textColor, border: borderStyle }}>
                    {(hasDiscussion || hasBirthday) && (
                      <span className="absolute top-0.5 flex items-center gap-0.5">
                        {hasDiscussion && <BookOpen size={8} style={{ color: '#D9C24C' }} />}
                        {hasBirthday && <Cake size={8} style={{ color: '#EFC94C' }} />}
                      </span>
                    )}
                    <span>{day}</span>
                    {(hasExempt || hasPersonal) && (
                      <span className="absolute bottom-0.5 flex items-center gap-0.5">
                        {hasExempt && <Plane size={8} style={{ color: INK }} />}
                        {hasPersonal && <User size={8} style={{ color: '#7FDCCF' }} />}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
            <div className="flex flex-wrap gap-2 mt-3">
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><Cake size={11} style={{ color: '#EFC94C' }} /> 생일</span>
              {DAY_TYPES.map((t) => t.key === '회식일' ? (
                <span key={t.key} className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: 'transparent', border: '1.5px solid rgba(229, 72, 77, 0.65)' }} /> {t.label}</span>
              ) : (
                <span key={t.key} className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: t.color }} /> {t.label}</span>
              ))}
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: WEEKEND_BG, border: `1px solid ${WEEKEND_TEXT}` }} /> 금·토·일(기본)</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><Plane size={11} /> 출장·휴가</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><User size={11} style={{ color: '#7FDCCF' }} /> 개인일정</span>
            </div>
            {(() => {
              const todayExcused = members.filter((m) => absenceExcuses.some((e) => e.date === todayStr() && e.member_id === m.id));
              if (todayExcused.length === 0) return null;
              return (
                <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="text-xs mb-1.5" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>오늘 참석불가</div>
                  <div className="flex flex-wrap gap-1.5">
                    {todayExcused.map((m) => {
                      const e = absenceExcuses.find((ee) => ee.date === todayStr() && ee.member_id === m.id);
                      return <span key={m.id} className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}><Stamp role={m.role} size={16} tilt={0} />{dispName(m.name, isLoggedIn)} · {e.reason}</span>;
                    })}
                  </div>
                </div>
              );
            })()}
            {selectedDate && (() => {
              const excusedMembers = members.filter((m) => absenceExcuses.some((e) => e.date === selectedDate && e.member_id === m.id));
              const birthdayFolksSelected = members.filter((m) => m.birthday && mdOf(m.birthday) === selectedDate.slice(5, 10));
              const myExcuse = currentMember ? absenceExcuses.find((e) => e.date === selectedDate && e.member_id === currentMember.id) : null;
              const setMyExcuseForDate = async (reason) => {
                if (!currentMember) return;
                await insertRow('absence_excuses', { id: uid('ae'), date: selectedDate, member_id: currentMember.id, reason });
                await reload();
              };
              const clearMyExcuseForDate = async () => {
                if (!myExcuse) return;
                await deleteRow('absence_excuses', 'id', myExcuse.id);
                await reload();
              };
              return (
                <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  {birthdayFolksSelected.length > 0 && (
                    <div className="mb-3">
                      <div className="text-xs mb-1.5 flex items-center gap-1" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}><Cake size={12} style={{ color: '#F0A87C' }} /> {fmtDate(selectedDate)} 생일</div>
                      <div className="flex flex-wrap gap-1.5">
                        {birthdayFolksSelected.map((m) => (
                          <span key={m.id} className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold" style={{ background: '#3A2E10', color: '#EFC94C' }}><Stamp role={m.role} size={16} tilt={0} />{dispName(m.name, isLoggedIn)}님</span>
                        ))}
                      </div>
                    </div>
                  )}
                  {excusedMembers.length > 0 && selectedDate !== todayStr() && (
                    <div className="mb-3">
                      <div className="text-xs mb-1.5" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(selectedDate)} 참석 불가</div>
                      <div className="flex flex-wrap gap-1.5">
                        {excusedMembers.map((m) => {
                          const e = absenceExcuses.find((ee) => ee.date === selectedDate && ee.member_id === m.id);
                          return <span key={m.id} className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}><Stamp role={m.role} size={16} tilt={0} />{dispName(m.name, isLoggedIn)} · {e.reason}</span>;
                        })}
                      </div>
                    </div>
                  )}
                  {currentMember && (
                    <div className="mb-3">
                      <div className="text-xs mb-1.5" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>본인 참석 불가 등록</div>
                      {myExcuse ? (
                        <div className="flex items-center gap-2">
                          <span className="text-xs rounded-full px-2 py-1" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>{myExcuse.reason}</span>
                          <button onClick={clearMyExcuseForDate} className="text-xs underline underline-offset-2" style={{ color: MUTE }}>취소</button>
                        </div>
                      ) : (
                        <div className="flex gap-1.5 flex-wrap">
                          <button onClick={() => setMyExcuseForDate('출장')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>출장</button>
                          <button onClick={() => setMyExcuseForDate('휴가')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>휴가</button>
                          <button onClick={() => setMyExcuseForDate('개인일정')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>개인일정</button>
                        </div>
                      )}
                    </div>
                  )}
                  {canManage && (
                    <div>
                      <div className="text-xs mb-2" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(selectedDate)} 유형 지정 (여러 개 선택 가능)</div>
                      <div className="flex flex-wrap gap-2">
                        {DAY_TYPES.map((t) => {
                          const active = getDayTypes(selectedDate).includes(t.key);
                          return (
                            <button key={t.key} onClick={() => toggleDayType(selectedDate, t.key)} className="rounded-full px-3 py-1.5 text-xs font-semibold border-2"
                              style={{ background: active ? t.bg : 'transparent', color: active ? t.color : MUTE, borderColor: active ? t.color : LINE }}>
                              {active && '✓ '}{t.label}
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  {!canManage && excusedMembers.length === 0 && <p className="text-xs" style={{ color: MUTE }}>이 날짜에 등록된 출장·휴가가 없어요.</p>}
                </div>
              );
            })()}
          </Card>

          {penaltyRule && (Object.values(penaltyByMember).some((p) => p.pending > 0) || warningMemberIds.size > 0) && (
            <Card><div className="flex items-center gap-1.5 text-sm font-semibold mb-1" style={{ color: INK }}><Gavel size={16} style={{ color: '#F0A87C' }} /> 벌칙 규정</div><p className="text-sm whitespace-pre-wrap" style={{ color: NEUTRAL_TEXT }}>{penaltyRule}</p></Card>
          )}

          <Card>
            <div className="text-sm font-semibold mb-2" style={{ color: INK }}>이번 달 출석률</div>
            <div className="flex flex-wrap items-center gap-2.5 mb-3">
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, background: '#7FDCCF' }} />출석</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, background: 'linear-gradient(90deg, #7FDCCF 50%, transparent 50%)', border: `1px solid ${LINE}` }} />절반출석(15~29분)</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, background: '#7FA8D9' }} />출장·휴가</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, border: `1px solid ${LINE}` }} />결석</span>
            </div>
            {weekChunkRanges.length > 0 && (
              <div className="flex items-center flex-wrap gap-x-3 mb-1.5" style={{ paddingLeft: 34 }}>
                {weekChunkRanges.map(([start, end], wi) => (
                  <div key={wi} className="text-center" style={{ width: (end - start) * 9 + (end - start - 1) * 4 }}>
                    <span className="text-[9px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{wi + 1}주</span>
                  </div>
                ))}
              </div>
            )}
            <div className="space-y-3">
              {rows.map((r, idx) => (
                <div key={r.id} className="space-y-1.5">
                  <div className="flex items-center justify-between gap-3">
                    <div className="flex items-center gap-2 min-w-0">
                      <span className="w-4 flex items-center justify-center shrink-0">
                        {r.rank <= 3 && r.present > 0 ? (
                          <Trophy size={14} color={r.rank === 1 ? '#EFC94C' : r.rank === 2 ? '#C9C9C9' : '#C08552'} strokeWidth={2.2} />
                        ) : (
                          <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.rank}</span>
                        )}
                      </span>
                      <Stamp role={r.role} size={24} tilt={0} /><span className="truncate" style={{ color: INK }}>{dispName(r.name, isLoggedIn)}</span>
                      {isCurrentMonth && warningMemberIds.has(r.id) && <span title="이번 주 열린 독서일을 지금까지 모두 결석 — 벌칙유의" className="inline-flex items-center gap-0.5 rounded-full px-1.5 py-0.5 text-[10px] font-semibold shrink-0" style={{ background: '#3A2213', color: '#F0A87C' }}>⚠️ 벌칙유의</span>}
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      {penaltyByMember[r.id]?.pending > 0 && <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}><Gavel size={10} /> 벌칙 대상 {penaltyByMember[r.id].pending}</span>}
                      <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.present}/{totalDays} · {r.rate}%</span>
                    </div>
                  </div>
                  <div className="flex items-center flex-wrap gap-y-1.5" style={{ paddingLeft: 34 }}>
                    {weekChunkRanges.map(([start, end], wi) => (
                      <React.Fragment key={wi}>
                        <div className="flex items-center gap-1">
                          {r.flags.slice(start, end).map((status, i) => (
                            <span key={i} className="rounded-full" style={{
                              width: 9, height: 9,
                              background: status === 'full' ? '#7FDCCF' : status === 'half' ? 'linear-gradient(90deg, #7FDCCF 50%, transparent 50%)' : status === 'excused' ? '#7FA8D9' : 'transparent',
                              border: status === 'full' || status === 'excused' ? 'none' : `1.5px solid ${LINE}`,
                            }} title={status === 'excused' ? '출장·휴가' : undefined} />
                          ))}
                        </div>
                        {wi < weekChunkRanges.length - 1 && (
                          <span className="shrink-0" style={{ width: 1, height: 11, background: MUTE, opacity: 0.4, transform: 'rotate(22deg)', margin: '0 7px' }} />
                        )}
                      </React.Fragment>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </Card>
        </>
      ) : (
        <>
          <Card>
            <div className="text-sm font-semibold mb-1" style={{ color: INK }}>전체 누적 출석률</div>
            <div className="text-xs mb-3" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>지금까지 총 출결 {totalSessions}회</div>
            <div className="grid grid-cols-3 gap-4">
              {allTimeRows.map((r, idx) => {
                const gaugeColor = r.rate >= 80 ? '#7FDCCF' : r.rate >= 50 ? '#EFC94C' : '#F0A87C';
                const medal = r.rank === 1 ? '🥇' : r.rank === 2 ? '🥈' : r.rank === 3 ? '🥉' : null;
                return (
                  <div key={r.id} className="flex flex-col items-center gap-1.5 text-center">
                    <span className="text-xs truncate max-w-full flex items-center justify-center gap-1" style={{ color: INK }}>
                      {medal && <span>{medal}</span>}{dispName(r.name, isLoggedIn)}
                    </span>
                    <div className="relative rounded-full shrink-0" style={{ width: 72, height: 72, background: `conic-gradient(${gaugeColor} ${r.rate * 3.6}deg, ${NEUTRAL_BG} ${r.rate * 3.6}deg 360deg)` }}>
                      <div className="absolute inset-[5px] rounded-full flex flex-col items-center justify-center" style={{ background: CARD_BG }}>
                        <span style={{ fontSize: 15, color: gaugeColor, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{r.rate}%</span>
                        <span style={{ fontSize: 10, color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.present}/{totalSessions}</span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

/* ---------------- 사용자관리 ---------------- */
function UsersScreen({ members, sortedMembers, currentUserId, setIdentity, canManage, notices, sessions, checkins, reload, requestDelete }) {
  const isLoggedIn = !!currentUserId;
  const [showAddForm, setShowAddForm] = useState(false);
  const [newName, setNewName] = useState(''); const [newRole, setNewRole] = useState('회원'); const [newBirthday, setNewBirthday] = useState(''); const [newPin, setNewPin] = useState('');
  const [newDept, setNewDept] = useState(''); const [newJobType, setNewJobType] = useState(''); const [newJoinedAt, setNewJoinedAt] = useState(''); const [newGenre, setNewGenre] = useState(''); const [newNote, setNewNote] = useState('');
  const [editingId, setEditingId] = useState(null);
  const [editMode, setEditMode] = useState('full'); // 'full' | 'self'
  const [editName, setEditName] = useState(''); const [editRole, setEditRole] = useState('회원'); const [editBirthday, setEditBirthday] = useState(''); const [editPin, setEditPin] = useState('');
  const [editDept, setEditDept] = useState(''); const [editJobType, setEditJobType] = useState(''); const [editJoinedAt, setEditJoinedAt] = useState(''); const [editGenre, setEditGenre] = useState(''); const [editNote, setEditNote] = useState('');

  const handleAdd = async () => {
    if (!newName.trim()) return;
    if (newPin && !/^\d{4}$/.test(newPin)) return;
    const id = uid('m');
    const isFirst = members.length === 0;
    await insertRow('members', { id, name: newName.trim(), role: newRole, birthday: newBirthday || null, pin: newPin || null, department: newDept || null, job_type: newJobType || null, joined_at: newJoinedAt || null, book_genre: newGenre || null, note: newNote || null });
    if (isFirst) setIdentity(id);
    await reload();
    setNewName(''); setNewRole('회원'); setNewBirthday(''); setNewPin(''); setNewDept(''); setNewJobType(''); setNewJoinedAt(''); setNewGenre(''); setNewNote(''); setShowAddForm(false);
  };
  const startEdit = (m) => { setEditingId(m.id); setEditMode('full'); setEditName(m.name); setEditRole(m.role); setEditBirthday(m.birthday || ''); setEditPin(m.pin || ''); setEditDept(m.department || ''); setEditJobType(m.job_type || ''); setEditJoinedAt(m.joined_at || ''); setEditGenre(m.book_genre || ''); setEditNote(m.note || ''); };
  const startSelfEdit = (m) => { setEditingId(m.id); setEditMode('self'); setEditBirthday(m.birthday || ''); setEditPin(m.pin || ''); setEditDept(m.department || ''); setEditJobType(m.job_type || ''); setEditJoinedAt(m.joined_at || ''); setEditGenre(m.book_genre || ''); setEditNote(m.note || ''); };
  const saveEdit = async () => {
    if (editPin && !/^\d{4}$/.test(editPin)) return;
    if (editMode === 'full') {
      if (!editName.trim()) return;
      await updateRow('members', 'id', editingId, { name: editName.trim(), role: editRole, birthday: editBirthday || null, pin: editPin || null, department: editDept || null, job_type: editJobType || null, joined_at: editJoinedAt || null, book_genre: editGenre || null, note: editNote || null });
    } else {
      await updateRow('members', 'id', editingId, { birthday: editBirthday || null, pin: editPin || null, department: editDept || null, job_type: editJobType || null, book_genre: editGenre || null });
    }
    await reload();
    setEditingId(null);
  };
  const removeMember = async (id) => { await deleteRow('members', 'id', id); await reload(); if (currentUserId === id) setIdentity(null); };

  const downloadExcel = () => {
    const data = sortedMembers.map((m) => ({ 이름: m.name, 직급: m.role, 소속: m.department || '', 직군: m.job_type || '', 생일: m.birthday || '', 가입일자: m.joined_at || '', 선호도서: m.book_genre || '', 비고: m.note || '', PIN: m.pin || '' }));
    const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '멤버명단'); XLSX.writeFile(wb, `책스초코_멤버명단_${todayStr()}.xlsx`);
  };
  const uploadExcel = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        const validRoles = ROLES.map((r) => r.key);
        const additions = rows.filter((r) => r['이름']).map((r) => ({
          id: uid('m'), name: String(r['이름']).trim(),
          role: validRoles.includes(r['구분']) ? r['구분'] : (validRoles.includes(r['직급']) ? r['직급'] : '회원'),
          birthday: /^\d{4}[.-]\d{2}[.-]\d{2}$/.test(r['생일'] || r['생년월일'] || '') ? String(r['생일'] || r['생년월일']).replace(/\./g, '-') : null,
          pin: /^\d{4}$/.test(String(r['PIN'] || '')) ? String(r['PIN']) : null,
          department: r['소속'] ? String(r['소속']) : null,
          job_type: r['직군'] ? String(r['직군']) : null,
          joined_at: r['가입일자'] ? String(r['가입일자']).replace(/\./g, '-') : null,
          book_genre: r['선호 도서 종류'] || r['선호도서'] ? String(r['선호 도서 종류'] || r['선호도서']) : null,
          note: r['비고'] ? String(r['비고']) : null,
        }));
        if (additions.length) { await supabase.from('members').insert(additions); await reload(); }
      } catch (err) {}
    };
    reader.readAsBinaryString(file); e.target.value = '';
  };

  return (
    <div className="space-y-4">
      {canManage && (
        <div className="flex gap-2">
          <GhostBtn onClick={downloadExcel} icon={Download}>명단 다운로드</GhostBtn>
          <label className="flex items-center justify-center gap-1.5 rounded-xl py-2.5 px-4 text-sm font-semibold cursor-pointer" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
            <Upload size={15} /> 엑셀 업로드<input type="file" accept=".xlsx,.xls" onChange={uploadExcel} className="hidden" />
          </label>
        </div>
      )}
      {!canManage && members.length > 0 && <div className="flex items-center gap-2 rounded-xl px-3.5 py-2.5 text-sm" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}><Lock size={15} /> 이름·직급 변경은 회장·간사·총무만 가능해요. 본인의 생일·PIN은 각자 수정할 수 있어요.</div>}

      <Card className="!p-0 overflow-hidden">
        {sortedMembers.length === 0 && <div className="px-4 py-8 text-center text-sm" style={{ color: MUTE }}>등록된 멤버가 없어요.</div>}
        {sortedMembers.map((m, idx) => (
          <div key={m.id} style={{ borderTop: idx === 0 ? 'none' : `1px solid ${ROW_LINE}` }}>
            {editingId === m.id ? (
              <div className="p-4 space-y-3" style={{ background: '#1A1812' }}>
                <div className="flex items-center gap-2 rounded-xl px-3 py-2" style={{ background: '#3A2E10' }}>
                  <Stamp role={m.role} size={28} tilt={0} />
                  <span className="font-semibold" style={{ color: '#EFC94C' }}>{m.name} 정보 수정</span>
                </div>
                {editMode === 'full' && (
                  <>
                    <input value={editName} onChange={(e) => setEditName(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                    <RolePicker value={editRole} onChange={setEditRole} />
                  </>
                )}
                {editMode === 'self' && <p className="text-xs" style={{ color: MUTE }}>이름·직급·가입일자·비고는 회장·간사·총무만 변경할 수 있어요. 나머지 정보는 본인이 직접 수정할 수 있어요.</p>}
                <div className="grid grid-cols-2 gap-2">
                  <input value={editDept} onChange={(e) => setEditDept(e.target.value)} placeholder="소속" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                  <input value={editJobType} onChange={(e) => setEditJobType(e.target.value)} placeholder="직군" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                </div>
                {editMode === 'full' && (
                  <>
                    <div><div className="text-xs mb-1" style={{ color: MUTE }}>가입일자</div><input type="date" value={editJoinedAt} onChange={(e) => setEditJoinedAt(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
                  </>
                )}
                <input value={editGenre} onChange={(e) => setEditGenre(e.target.value)} placeholder="선호 도서 종류 (예: 소설, 자기계발)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                {editMode === 'full' && (
                  <input value={editNote} onChange={(e) => setEditNote(e.target.value)} placeholder="비고" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                )}
                <div><div className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Cake size={13} /> 생일</div><input type="date" value={editBirthday} onChange={(e) => setEditBirthday(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
                <div><div className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Lock size={13} /> 본인 확인 PIN (4자리, 선택)</div><input type="password" inputMode="numeric" maxLength={4} value={editPin} onChange={(e) => setEditPin(e.target.value.replace(/\D/g, ''))} className="w-full rounded-xl border px-3 py-2 text-sm tracking-[0.3em] outline-none" style={inputStyle} placeholder="설정 안 함" /></div>
                <div className="flex gap-2 pt-1"><PrimaryBtn onClick={saveEdit} icon={Check}>저장</PrimaryBtn><GhostBtn onClick={() => setEditingId(null)} icon={X}>취소</GhostBtn></div>
              </div>
            ) : (
              <div className="flex items-center justify-between px-4 py-3">
                <div className="flex items-center gap-3 min-w-0 flex-1">
                  <Stamp role={m.role} size={36} tilt={idx % 2 === 0 ? -5 : 4} />
                  <div className="min-w-0 flex-1">
                    <div className="font-semibold truncate" style={{ color: INK }}>{dispName(m.name, isLoggedIn)}{m.id === currentUserId && <span className="ml-1.5 text-[11px] font-normal" style={{ color: MUTE }}>(나)</span>}</div>
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <RoleChip role={m.role} />
                      {m.birthday && isLoggedIn && <span className="text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtMD(mdOf(m.birthday))}</span>}
                      {m.pin && <Lock size={11} style={{ color: MUTE }} />}
                    </div>
                    {isLoggedIn && (m.department || m.job_type || m.joined_at || m.book_genre || m.note) && (
                      <div className="flex items-center gap-2 flex-wrap mt-1 text-[11px]" style={{ color: MUTE }}>
                        {(m.department || m.job_type) && (
                          <span>{m.department}{m.department && m.job_type ? `(${m.job_type})` : m.job_type}</span>
                        )}
                        {m.joined_at && <span>가입 {fmtDate(m.joined_at)}</span>}
                        {m.book_genre && <span>{m.book_genre}</span>}
                        {m.note && <span className="italic">{m.note}</span>}
                      </div>
                    )}
                  </div>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  {canManage && <button onClick={() => startEdit(m)} className="p-2 rounded-lg" style={{ color: MUTE }}><Pencil size={16} /></button>}
                  {!canManage && m.id === currentUserId && <button onClick={() => startSelfEdit(m)} className="p-2 rounded-lg" style={{ color: MUTE }}><Pencil size={16} /></button>}
                  {canManage && <button onClick={() => requestDelete(() => removeMember(m.id), `${m.name}님을 삭제할까요? 관련 기록도 함께 사라져요.`)} className="p-2 rounded-lg" style={{ color: '#F0A87C' }}><Trash2 size={16} /></button>}
                </div>
              </div>
            )}
          </div>
        ))}
      </Card>

      {(canManage || sortedMembers.length === 0) && (
        showAddForm ? (
          <Card className="space-y-3">
            <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="이름을 입력하세요" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <RolePicker value={newRole} onChange={setNewRole} />
            <div className="grid grid-cols-2 gap-2">
              <input value={newDept} onChange={(e) => setNewDept(e.target.value)} placeholder="소속 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
              <input value={newJobType} onChange={(e) => setNewJobType(e.target.value)} placeholder="직군 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            </div>
            <div><div className="text-xs mb-1" style={{ color: MUTE }}>가입일자 (선택)</div><input type="date" value={newJoinedAt} onChange={(e) => setNewJoinedAt(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            <input value={newGenre} onChange={(e) => setNewGenre(e.target.value)} placeholder="선호 도서 종류 (선택)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <input value={newNote} onChange={(e) => setNewNote(e.target.value)} placeholder="비고 (선택)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <div><div className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Cake size={13} /> 생일 (선택)</div><input type="date" value={newBirthday} onChange={(e) => setNewBirthday(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            <div><div className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Lock size={13} /> 본인 확인 PIN (4자리, 선택)</div><input type="password" inputMode="numeric" maxLength={4} value={newPin} onChange={(e) => setNewPin(e.target.value.replace(/\D/g, ''))} className="w-full rounded-xl border px-3 py-2 text-sm tracking-[0.3em] outline-none" style={inputStyle} placeholder="설정 안 함" /></div>
            <div className="flex gap-2"><PrimaryBtn onClick={handleAdd} icon={Check}>등록</PrimaryBtn><GhostBtn onClick={() => setShowAddForm(false)} icon={X}>취소</GhostBtn></div>
          </Card>
        ) : (
          <button onClick={() => setShowAddForm(true)} className="w-full flex items-center justify-center gap-1.5 rounded-2xl border-2 border-dashed py-3 text-sm font-semibold" style={{ borderColor: LINE, color: MUTE }}><Plus size={16} /> 멤버 추가</button>
        )
      )}
    </div>
  );
}

/* ---------------- 출석관리 (간사 전용) ---------------- */
/* ---------------- 회계 (총무 관리) ---------------- */
function TreasuryScreen({ members, duesPayments, expenses, dinnerCollections, currentMember, reload, requestDelete, showToast }) {
  const [cursor, setCursor] = useState(new Date());
  const monthKey = `${cursor.getFullYear()}-${pad(cursor.getMonth() + 1)}`;
  const shift = (delta) => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + delta, 1));

  // 카드 접기/펼치기
  const [duesOpen, setDuesOpen] = useState(true);
  const [dinnerOpen, setDinnerOpen] = useState(true);
  const [expenseOpen, setExpenseOpen] = useState(true);
  const SectionHeader = ({ title, open, onToggle, right }) => (
    <div className="flex items-center justify-between">
      <button onClick={onToggle} className="flex items-center gap-1.5">
        {open ? <ChevronUp size={15} style={{ color: MUTE }} /> : <ChevronDown size={15} style={{ color: MUTE }} />}
        <span className="text-sm font-semibold" style={{ color: INK }}>{title}</span>
      </button>
      {right}
    </div>
  );

  const [defaultAmount, setDefaultAmount] = useState('5000');
  const duesForMonth = duesPayments.filter((d) => d.month === monthKey);
  const getDues = (memberId) => duesForMonth.find((d) => d.member_id === memberId);
  const togglePaid = async (memberId) => {
    const existing = getDues(memberId);
    if (existing) {
      await updateRow('dues_payments', 'id', existing.id, { paid: !existing.paid, paid_at: !existing.paid ? new Date().toISOString() : null });
    } else {
      const amt = parseInt(defaultAmount, 10) || 0;
      await insertRow('dues_payments', { id: uid('dp'), member_id: memberId, month: monthKey, amount: amt, paid: true, paid_at: new Date().toISOString() });
    }
    await reload();
  };
  const totalDuesThisMonth = duesForMonth.filter((d) => d.paid).reduce((sum, d) => sum + Number(d.amount), 0);

  // 개별 금액 수정 (완납/미납과 무관하게 각 멤버 회비 금액을 조정)
  const [amountEdits, setAmountEdits] = useState({});
  const getAmountValue = (m) => {
    if (amountEdits[m.id] !== undefined) return amountEdits[m.id];
    const d = getDues(m.id);
    return d ? String(d.amount) : defaultAmount;
  };
  const saveAmount = async (memberId) => {
    const val = amountEdits[memberId];
    if (val === undefined) return;
    const amt = parseInt(val, 10) || 0;
    const existing = getDues(memberId);
    if (existing) { if (Number(existing.amount) !== amt) await updateRow('dues_payments', 'id', existing.id, { amount: amt }); }
    else await insertRow('dues_payments', { id: uid('dp'), member_id: memberId, month: monthKey, amount: amt, paid: false, paid_at: null });
    await reload();
    setAmountEdits((prev) => { const next = { ...prev }; delete next[memberId]; return next; });
  };

  // 일괄 납부 처리 — 이번 달 미납 회원 전체를 각자 현재 금액(또는 기본 금액)으로 완납 처리
  const unpaidCount = members.filter((m) => !getDues(m.id)?.paid).length;

  // 멤버별 미납 회식비 합계 (전체 기간, 회비와는 별개로 회비 카드에 참고 표시)
  const unpaidDinnerByMember = {};
  dinnerCollections.forEach((c) => { if (!c.paid) unpaidDinnerByMember[c.member_id] = (unpaidDinnerByMember[c.member_id] || 0) + Number(c.amount); });
  const bulkPayAll = async () => {
    for (const m of members) {
      const existing = getDues(m.id);
      if (existing?.paid) continue;
      const amt = existing ? Number(existing.amount) : (parseInt(defaultAmount, 10) || 0);
      if (existing) await updateRow('dues_payments', 'id', existing.id, { paid: true, paid_at: new Date().toISOString() });
      else await insertRow('dues_payments', { id: uid('dp'), member_id: m.id, month: monthKey, amount: amt, paid: true, paid_at: new Date().toISOString() });
    }
    await reload();
    showToast?.('일괄 납부 처리했어요.', 'success');
  };

  const [expDate, setExpDate] = useState(todayStr());
  const [expDesc, setExpDesc] = useState('');
  const [expAmount, setExpAmount] = useState('');
  const expensesForMonth = expenses.filter((e) => e.date.startsWith(monthKey)).sort((a, b) => b.date.localeCompare(a.date));
  const totalExpensesThisMonth = expensesForMonth.reduce((sum, e) => sum + Number(e.amount), 0);
  const addExpense = async () => {
    if (!expDesc.trim() || !expAmount) return;
    await insertRow('expenses', { id: uid('ex'), date: expDate, description: expDesc.trim(), amount: parseInt(expAmount, 10) || 0, recorded_by: currentMember?.name || '', created_at: new Date().toISOString() });
    await reload();
    setExpDesc(''); setExpAmount('');
    showToast?.('지출 내역을 등록했어요.', 'success');
  };
  const removeExpense = async (id) => { await deleteRow('expenses', 'id', id); await reload(); };

  // 회식비 정산 — 특정 날짜에 1차/2차/... 금액을 발생할 때마다 등록 (식당명은 "회식 N차 · 식당명" 형태로 저장)
  const [dinnerDate, setDinnerDate] = useState(todayStr());
  const [dinnerAmount, setDinnerAmount] = useState('');
  const dinnerRoundRe = /^회식 (\d+)차(?: · (.*))?$/;
  const dinnerExpenses = expenses.filter((e) => e.date === dinnerDate && dinnerRoundRe.test(e.description))
    .sort((a, b) => parseInt(a.description.match(dinnerRoundRe)[1], 10) - parseInt(b.description.match(dinnerRoundRe)[1], 10));
  const nextDinnerRound = dinnerExpenses.length + 1;
  const dinnerTotal = dinnerExpenses.reduce((sum, e) => sum + Number(e.amount), 0);
  const addDinnerRound = async () => {
    if (!dinnerAmount) return;
    await insertRow('expenses', { id: uid('ex'), date: dinnerDate, description: `회식 ${nextDinnerRound}차`, amount: parseInt(dinnerAmount, 10) || 0, recorded_by: currentMember?.name || '', created_at: new Date().toISOString() });
    await reload();
    setDinnerAmount('');
    showToast?.(`회식 ${nextDinnerRound}차를 등록했어요.`, 'success');
  };

  // 식당명 수정
  const [restaurantEdits, setRestaurantEdits] = useState({});
  const getRestaurantValue = (e) => {
    if (restaurantEdits[e.id] !== undefined) return restaurantEdits[e.id];
    const m = e.description.match(dinnerRoundRe);
    return m && m[2] ? m[2] : '';
  };
  const saveRestaurant = async (e) => {
    const val = restaurantEdits[e.id];
    if (val === undefined) return;
    const m = e.description.match(dinnerRoundRe);
    const roundNum = m ? m[1] : '';
    const newDesc = `회식 ${roundNum}차` + (val.trim() ? ` · ${val.trim()}` : '');
    if (newDesc !== e.description) await updateRow('expenses', 'id', e.id, { description: newDesc });
    await reload();
    setRestaurantEdits((prev) => { const next = { ...prev }; delete next[e.id]; return next; });
  };

  // 회식 정산 계산기 — 차수별로 정산 방식(회비/각출)과 참석자를 선택해 1인당 징수액 계산
  const SETTLE_MODES = [
    { key: 'club', label: '회비 차감 후 정산' },
    { key: 'split', label: '각출 정산' },
  ];
  const [roundSettlement, setRoundSettlement] = useState({});
  const getRoundSettlement = (roundId) => roundSettlement[roundId] || { mode: 'club', attendees: [], headcount: '' };
  const setRoundMode = (roundId, mode) => setRoundSettlement((prev) => ({ ...prev, [roundId]: { ...getRoundSettlement(roundId), mode } }));
  const setRoundHeadcount = (roundId, headcount) => setRoundSettlement((prev) => ({ ...prev, [roundId]: { ...getRoundSettlement(roundId), headcount } }));
  const toggleRoundAttendee = (roundId, memberId) => setRoundSettlement((prev) => {
    const cur = getRoundSettlement(roundId);
    const attendees = cur.attendees.includes(memberId) ? cur.attendees.filter((id) => id !== memberId) : [...cur.attendees, memberId];
    return { ...prev, [roundId]: { ...cur, attendees } };
  });

  const totalDuesAllTime = duesPayments.filter((d) => d.paid).reduce((sum, d) => sum + Number(d.amount), 0);
  const totalCollectionsAllTime = dinnerCollections.filter((c) => c.paid).reduce((sum, c) => sum + Number(c.amount), 0);
  const totalExpensesAllTime = expenses.reduce((sum, e) => sum + Number(e.amount), 0);
  const balance = totalDuesAllTime + totalCollectionsAllTime - totalExpensesAllTime;
  const fmtWon = (n) => `${Math.round(n).toLocaleString('ko-KR')}원`;

  // 차수별 정산 결과: 현재 잔액(balance)은 이미 이 차수 지출이 반영된 값이라고 보고,
  // 이 차수를 빼기 전 잔액(preRoundBalance)을 기준으로 부족액(shortfall)을 판단한다.
  const computeRoundSettlement = (e) => {
    const rs = getRoundSettlement(e.id);
    const roundCost = Number(e.amount);
    const preRoundBalance = balance + roundCost;
    const shortfall = Math.max(0, roundCost - preRoundBalance); // 회비만으로 감당 안 되는 금액
    let count = 0;
    let collection = 0;
    if (rs.mode === 'split') {
      count = rs.attendees.length;
      collection = roundCost;
    } else if (shortfall > 0) {
      count = rs.attendees.length > 0 ? rs.attendees.length : (parseInt(rs.headcount, 10) || 0);
      collection = shortfall;
    }
    const perPerson = count > 0 ? collection / count : 0;
    const finalBalance = balance + collection;
    return { ...rs, roundCost, shortfall, count, collection, perPerson, finalBalance };
  };
  const dinnerSettlements = dinnerExpenses.map((e) => ({ id: e.id, ...computeRoundSettlement(e) }));
  const totalDinnerCollection = dinnerSettlements.reduce((sum, s) => sum + s.collection, 0);
  const finalBalanceAfterDinner = balance + totalDinnerCollection;

  // 차수별 징수 내역(개인별 배분·완납 여부) — N분의 1 나머지는 참석 순서상 앞사람부터 1원씩 더 배분
  const getRoundCollections = (roundId) => dinnerCollections.filter((c) => c.expense_id === roundId);
  const generateCollections = async (e, s, opts = {}) => {
    if (s.collection <= 0 || s.attendees.length === 0) return;
    const orderedIds = members.filter((m) => s.attendees.includes(m.id)).map((m) => m.id);
    const base = Math.floor(s.collection / orderedIds.length);
    const remainder = s.collection - base * orderedIds.length;
    const presidentId = members.find((m) => m.role === '회장' && s.attendees.includes(m.id))?.id;
    const amountFor = (mid, idx) => {
      if (presidentId) return base + (mid === presidentId ? remainder : 0);
      return base + (idx < remainder ? 1 : 0); // 참석자 중 회장이 없으면 참석 순서상 앞사람이 나머지를 부담
    };
    const existing = getRoundCollections(e.id);
    const existingMap = {}; existing.forEach((c) => { existingMap[c.member_id] = c; });
    for (const c of existing) { if (!orderedIds.includes(c.member_id)) await deleteRow('dinner_collections', 'id', c.id); }
    for (let i = 0; i < orderedIds.length; i++) {
      const mid = orderedIds[i];
      const amt = amountFor(mid, i);
      const prev = existingMap[mid];
      if (prev) { if (Number(prev.amount) !== amt) await updateRow('dinner_collections', 'id', prev.id, { amount: amt }); }
      else await insertRow('dinner_collections', { id: uid('dc'), expense_id: e.id, member_id: mid, amount: amt, paid: false, paid_at: null });
    }
    if (!opts.skipReload) await reload();
  };
  const generateAllCollections = async () => {
    for (const s of dinnerSettlements) {
      const e = dinnerExpenses.find((x) => x.id === s.id);
      if (e) await generateCollections(e, s, { skipReload: true });
    }
    await reload();
  };
  // 한 멤버가 여러 차수에 걸쳐 낼 금액을 한 번에 완납/미납 처리
  const toggleMemberAllPaid = async (memberId) => {
    const memberCollections = dinnerExpenses.flatMap((e) => getRoundCollections(e.id)).filter((c) => c.member_id === memberId);
    if (memberCollections.length === 0) return;
    const allPaid = memberCollections.every((c) => c.paid);
    const newPaid = !allPaid;
    for (const c of memberCollections) { if (c.paid !== newPaid) await updateRow('dinner_collections', 'id', c.id, { paid: newPaid, paid_at: newPaid ? new Date().toISOString() : null }); }
    await reload();
  };

  const [dinnerActionMsg, setDinnerActionMsg] = useState('');
  const handleGenerateAll = async () => {
    const targets = dinnerSettlements.filter((s) => s.collection > 0 && s.attendees.length > 0);
    if (targets.length === 0) { setDinnerActionMsg('참석자가 선택된 차수가 없어요. 각 차수에서 정산 방식을 정하고 참석자를 선택해 주세요.'); return; }
    try {
      await generateAllCollections();
      setDinnerActionMsg('');
      showToast?.('최종 부담액을 생성했어요.', 'success');
    } catch (err) {
      console.error(err);
      setDinnerActionMsg('부담금 생성에 실패했어요. Supabase에 dinner_collections 테이블이 만들어져 있는지 확인해 주세요.');
      showToast?.('부담금 생성에 실패했어요.', 'error');
    }
  };

  // 회식 최종 정산 명단 — 선택한 날짜의 모든 차수에서 생성된 개인별 징수 내역을 합산
  const allDinnerCollectionsForDate = dinnerExpenses.flatMap((e) => getRoundCollections(e.id));
  const finalMemberTotalsMap = {};
  allDinnerCollectionsForDate.forEach((c) => {
    if (!finalMemberTotalsMap[c.member_id]) finalMemberTotalsMap[c.member_id] = { amount: 0, paid: true };
    finalMemberTotalsMap[c.member_id].amount += Number(c.amount);
    if (!c.paid) finalMemberTotalsMap[c.member_id].paid = false;
  });
  const finalMemberTotals = members
    .filter((m) => finalMemberTotalsMap[m.id])
    .map((m) => ({ member: m, ...finalMemberTotalsMap[m.id] }))
    .sort((a, b) => b.amount - a.amount);

  // 미납자 명단 공유 (회비 + 회식비 미납 내역 함께) — 모바일에서는 공유 시트, 아니면 클립보드 복사
  const [duesCopied, setDuesCopied] = useState(false);
  const copyUnpaidList = async () => {
    const unpaidNames = members.filter((m) => !getDues(m.id)?.paid).map((m) => m.name);
    const duesLine = unpaidNames.length ? `[회비 미납]\n${unpaidNames.join(', ')}` : '[회비 미납] 없음 🎉';
    const dinnerLines = members
      .filter((m) => unpaidDinnerByMember[m.id] > 0)
      .map((m) => `${m.name} ${fmtWon(unpaidDinnerByMember[m.id])}`);
    const dinnerBlock = dinnerLines.length ? `\n\n[회식비 미납]\n${dinnerLines.join('\n')}` : '';
    const text = `[${cursor.getFullYear()}.${cursor.getMonth() + 1}]\n${duesLine}${dinnerBlock}`;
    if (navigator.share) {
      try { await navigator.share({ text }); return; } catch (err) { /* 공유 취소 시 아무 것도 안 함 */ return; }
    }
    try { await navigator.clipboard.writeText(text); setDuesCopied(true); setTimeout(() => setDuesCopied(false), 2000); showToast?.('미납자 명단을 복사했어요.', 'success'); } catch (err) { showToast?.('복사에 실패했어요.', 'error'); }
  };

  // 월별 수입·지출 추이 (최근 6개월)
  const trendMonths = Array.from({ length: 6 }).map((_, i) => { const d = new Date(cursor.getFullYear(), cursor.getMonth() - 5 + i, 1); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`; });
  const trendStats = trendMonths.map((mk) => {
    const income = duesPayments.filter((d) => d.paid && d.month === mk).reduce((sum, d) => sum + Number(d.amount), 0)
      + dinnerCollections.filter((c) => c.paid && expenses.find((e) => e.id === c.expense_id)?.date.startsWith(mk)).reduce((sum, c) => sum + Number(c.amount), 0);
    const expense = expenses.filter((e) => e.date.startsWith(mk)).reduce((sum, e) => sum + Number(e.amount), 0);
    return { mk, income, expense };
  });
  const trendMax = Math.max(1, ...trendStats.flatMap((t) => [t.income, t.expense]));

  return (
    <div className="space-y-4">
      <Card>
        <div className="flex items-center justify-between mb-1">
          <button onClick={() => shift(-1)} className="p-1.5" style={{ color: MUTE }}><ChevronLeft size={18} /></button>
          <div className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>{cursor.getFullYear()}년 {cursor.getMonth() + 1}월</div>
          <button onClick={() => shift(1)} className="p-1.5" style={{ color: MUTE }}><ChevronRight size={18} /></button>
        </div>
      </Card>

      <Card>
        <div className="text-sm font-semibold mb-1" style={{ color: INK }}>종합 장부</div>
        <div className="grid grid-cols-2 gap-2 mt-2 text-center">
          <div className="rounded-xl py-2" style={{ background: '#12302C' }}>
            <div className="text-[11px]" style={{ color: MUTE }}>이번 달 회비 수입</div>
            <div className="font-semibold" style={{ color: '#7FDCCF' }}>{fmtWon(totalDuesThisMonth)}</div>
          </div>
          <div className="rounded-xl py-2" style={{ background: '#3A2213' }}>
            <div className="text-[11px]" style={{ color: MUTE }}>이번 달 지출</div>
            <div className="font-semibold" style={{ color: '#F0A87C' }}>{fmtWon(totalExpensesThisMonth)}</div>
          </div>
        </div>
        <div className="mt-2 pt-2 flex items-center justify-between" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
          <span className="text-xs" style={{ color: MUTE }}>전체 누적 잔액</span>
          <span className="font-semibold" style={{ color: balance >= 0 ? '#7FDCCF' : '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtWon(balance)}</span>
        </div>
      </Card>

      <Card>
        <div className="text-sm font-semibold mb-3" style={{ color: INK }}>월별 수입·지출 추이 (최근 6개월)</div>
        <div className="flex items-end justify-between gap-1.5" style={{ height: 96 }}>
          {trendStats.map((t) => (
            <div key={t.mk} className="flex-1 flex items-end justify-center gap-1" style={{ height: '100%' }}>
              <div className="rounded-t-sm" style={{ width: 9, height: `${Math.max(2, (t.income / trendMax) * 100)}%`, background: '#7FDCCF' }} title={`수입 ${fmtWon(t.income)}`} />
              <div className="rounded-t-sm" style={{ width: 9, height: `${Math.max(2, (t.expense / trendMax) * 100)}%`, background: '#F0A87C' }} title={`지출 ${fmtWon(t.expense)}`} />
            </div>
          ))}
        </div>
        <div className="flex items-center justify-between mt-1.5">
          {trendStats.map((t) => (
            <span key={t.mk} className="flex-1 text-center text-[10px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{t.mk.slice(5)}월</span>
          ))}
        </div>
        <div className="flex items-center gap-3 mt-2 pt-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
          <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="inline-block rounded-sm" style={{ width: 8, height: 8, background: '#7FDCCF' }} /> 수입</span>
          <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="inline-block rounded-sm" style={{ width: 8, height: 8, background: '#F0A87C' }} /> 지출</span>
        </div>
      </Card>

      <Card>
        <SectionHeader title="회비 납부 현황" open={duesOpen} onToggle={() => setDuesOpen((v) => !v)} right={
          <div className="flex items-center gap-1.5">
            <span className="text-xs" style={{ color: MUTE }}>기본 금액</span>
            <input type="number" value={defaultAmount} onChange={(e) => setDefaultAmount(e.target.value)} className="w-20 rounded-lg border px-2 py-1 text-xs outline-none" style={inputStyle} />
          </div>
        } />
        {duesOpen && (
          <>
        <div className="space-y-1.5 mt-2">
          {members.map((m) => {
            const d = getDues(m.id);
            const paid = d?.paid;
            return (
              <div key={m.id} className="flex items-center justify-between py-1.5" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                <div className="flex items-center gap-2 min-w-0">
                  <Stamp role={m.role} size={24} tilt={0} />
                  <div className="min-w-0">
                    <span className="text-sm truncate" style={{ color: INK }}>{m.name}</span>
                    {unpaidDinnerByMember[m.id] > 0 && <div className="text-[10px]" style={{ color: '#F0A87C' }}>회식비 미납 {fmtWon(unpaidDinnerByMember[m.id])}</div>}
                  </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <input type="number" value={getAmountValue(m)} onChange={(e) => setAmountEdits((prev) => ({ ...prev, [m.id]: e.target.value }))} onBlur={() => saveAmount(m.id)} className="w-20 rounded-lg border px-2 py-1 text-xs outline-none text-right" style={inputStyle} />
                  <button onClick={() => togglePaid(m.id)} className="inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-semibold" style={{ background: paid ? '#12302C' : NEUTRAL_BG, color: paid ? '#7FDCCF' : MUTE }}>{paid ? <Check size={12} /> : null} {paid ? '완납' : '미납'}</button>
                </div>
              </div>
            );
          })}
        </div>
        <div className="pt-3 mt-1 space-y-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
          <PrimaryBtn onClick={bulkPayAll} disabled={unpaidCount === 0} icon={Check}>미납 {unpaidCount}명 일괄 납부 처리</PrimaryBtn>
          <button onClick={copyUnpaidList} className="w-full rounded-xl py-2 text-xs font-semibold" style={{ background: NEUTRAL_BG, color: duesCopied ? '#7FDCCF' : MUTE }}>{duesCopied ? '복사했어요 ✓' : '미납자 명단 복사'}</button>
        </div>
          </>
        )}
      </Card>

      <Card className="space-y-2">
        <SectionHeader title="회식비 정산" open={dinnerOpen} onToggle={() => setDinnerOpen((v) => !v)} right={
          <input type="date" value={dinnerDate} onChange={(e) => setDinnerDate(e.target.value)} className="rounded-lg border px-2 py-1 text-xs outline-none" style={inputStyle} />
        } />
        {dinnerOpen && (
          <>
        {dinnerExpenses.length > 0 && (
          <div className="space-y-2" style={{ borderTop: `1px solid ${ROW_LINE}`, paddingTop: 8 }}>
            {dinnerSettlements.map((s) => {
              const e = dinnerExpenses.find((x) => x.id === s.id);
              return (
                <div key={s.id} className="rounded-xl p-2.5 space-y-2" style={{ background: NEUTRAL_BG }}>
                  <div className="flex items-center justify-between text-sm">
                    <span className="font-semibold" style={{ color: INK }}>회식 {e.description.match(dinnerRoundRe)?.[1]}차</span>
                    <div className="flex items-center gap-2">
                      <span style={{ color: '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtWon(s.roundCost)}</span>
                      <button onClick={() => requestDelete(() => removeExpense(e.id), '이 지출 내역을 삭제할까요?')} className="p-1" style={{ color: MUTE }}><Trash2 size={14} /></button>
                    </div>
                  </div>
                  <input value={getRestaurantValue(e)} onChange={(ev) => setRestaurantEdits((prev) => ({ ...prev, [e.id]: ev.target.value }))} onBlur={() => saveRestaurant(e)} placeholder="식당명" className="w-full rounded-lg border px-2 py-1.5 text-xs outline-none" style={inputStyle} />
                  <div className="grid grid-cols-2 gap-2">
                    {SETTLE_MODES.map((sm) => (
                      <button key={sm.key} onClick={() => setRoundMode(s.id, sm.key)} className="rounded-xl px-3 py-2.5 text-sm font-semibold border"
                        style={{ background: s.mode === sm.key ? '#3A2E10' : CARD_BG, color: s.mode === sm.key ? '#EFC94C' : MUTE, borderColor: s.mode === sm.key ? '#EFC94C' : LINE }}>{sm.label}</button>
                    ))}
                  </div>
                  {s.mode === 'split' && (
                    <div className="flex flex-wrap gap-1.5">
                      {members.map((m) => {
                        const checked = s.attendees.includes(m.id);
                        return (
                          <button key={m.id} onClick={() => toggleRoundAttendee(s.id, m.id)}
                            className="flex items-center gap-1 rounded-full border pl-1 pr-2 py-0.5"
                            style={{ borderColor: checked ? '#7FA8D9' : LINE, background: checked ? '#1E2A38' : 'transparent' }}>
                            <Stamp role={m.role} size={16} tilt={0} />
                            <span className="text-[11px]" style={{ color: checked ? '#7FA8D9' : INK }}>{m.name}</span>
                          </button>
                        );
                      })}
                    </div>
                  )}
                  {s.mode === 'club' && s.shortfall > 0 && (
                    <div className="space-y-1.5">
                      <p className="text-[11px]" style={{ color: '#F0A87C' }}>⚠ 회비 잔액이 부족해요 ({fmtWon(s.shortfall)} 부족). 참석자를 선택하거나 인원수를 입력하면 각출액을 계산해요.</p>
                      <div className="flex flex-wrap gap-1.5">
                        {members.map((m) => {
                          const checked = s.attendees.includes(m.id);
                          return (
                            <button key={m.id} onClick={() => toggleRoundAttendee(s.id, m.id)}
                              className="flex items-center gap-1 rounded-full border pl-1 pr-2 py-0.5"
                              style={{ borderColor: checked ? '#7FA8D9' : LINE, background: checked ? '#1E2A38' : 'transparent' }}>
                              <Stamp role={m.role} size={16} tilt={0} />
                              <span className="text-[11px]" style={{ color: checked ? '#7FA8D9' : INK }}>{m.name}</span>
                            </button>
                          );
                        })}
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="text-[11px] shrink-0" style={{ color: MUTE }}>또는 인원수</span>
                        <input type="number" value={s.headcount} onChange={(ev) => setRoundHeadcount(s.id, ev.target.value)} placeholder="인원수" className="w-24 rounded-lg border px-2 py-1 text-xs outline-none" style={inputStyle} />
                      </div>
                    </div>
                  )}
                  <div className="pt-1 space-y-1" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                    <div className="flex items-center justify-between text-xs">
                      <span style={{ color: MUTE }}>참석 {s.count}명 · 1인당 평균 부담액</span>
                      <span className="font-semibold" style={{ color: '#EFC94C' }}>{fmtWon(s.perPerson)}</span>
                    </div>
                    <div className="flex items-center justify-between text-xs">
                      <span style={{ color: MUTE }}>이 차수 정산 후 예상 잔액</span>
                      <span className="font-semibold" style={{ color: s.finalBalance >= 0 ? '#7FDCCF' : '#F0A87C' }}>{fmtWon(s.finalBalance)}</span>
                    </div>
                  </div>
                </div>
              );
            })}
            <div className="pt-1 space-y-1">
              <div className="flex items-center justify-between text-xs" style={{ color: MUTE }}>
                <span>{fmtDate(dinnerDate)} 회식비 합계</span>
                <span className="font-semibold" style={{ color: '#F0A87C' }}>{fmtWon(dinnerTotal)}</span>
              </div>
              <div className="flex items-center justify-between text-sm">
                <span style={{ color: MUTE }}>회식 정산 후 최종 결론 잔액</span>
                <span className="font-semibold" style={{ color: finalBalanceAfterDinner >= 0 ? '#7FDCCF' : '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtWon(finalBalanceAfterDinner)}</span>
              </div>
            </div>
            <div className="pt-2 flex justify-center">
              <button onClick={handleGenerateAll} className="w-full max-w-xs rounded-xl py-3 text-sm font-semibold text-center" style={{ background: BTN_BG, color: BTN_TEXT }}>최종 부담액 생성</button>
            </div>
            {dinnerActionMsg && <p className="text-xs text-center" style={{ color: '#F0A87C' }}>{dinnerActionMsg}</p>}
            {finalMemberTotals.length > 0 && (
              <div className="pt-2 mt-1" style={{ borderTop: `1px dashed ${ROW_LINE}` }}>
                <div className="flex items-center gap-1.5 mb-2">
                  <span className="text-xs font-semibold" style={{ color: INK }}>회식 최종 정산표 (차수별 부담액 · 합계)</span>
                  {finalMemberTotals.every((f) => f.paid) && <span className="rounded-full px-2 py-0.5 text-[10px] font-semibold" style={{ background: '#12302C', color: '#7FDCCF' }}>정산 완료 ✓</span>}
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs" style={{ borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ borderBottom: `1px solid ${ROW_LINE}` }}>
                        <th className="text-left pb-1.5 pr-2 font-medium" style={{ color: MUTE }}>이름</th>
                        {dinnerExpenses.map((e) => (
                          <th key={e.id} className="text-right pb-1.5 px-2 font-medium whitespace-nowrap" style={{ color: MUTE }}>{e.description.match(dinnerRoundRe)?.[1]}차</th>
                        ))}
                        <th className="text-right pb-1.5 pl-2 font-medium" style={{ color: MUTE }}>합계</th>
                        <th className="text-center pb-1.5 pl-2 font-medium" style={{ color: MUTE }}>완납</th>
                      </tr>
                    </thead>
                    <tbody>
                      {finalMemberTotals.map(({ member: m, amount, paid }) => (
                        <tr key={m.id} style={{ borderBottom: `1px solid ${ROW_LINE}` }}>
                          <td className="py-1.5 pr-2 whitespace-nowrap" style={{ color: INK }}>{m.name}</td>
                          {dinnerExpenses.map((e) => {
                            const c = getRoundCollections(e.id).find((cc) => cc.member_id === m.id);
                            return <td key={e.id} className="text-right py-1.5 px-2" style={{ color: c ? MUTE : LINE, fontFamily: "'IBM Plex Mono', monospace" }}>{c ? fmtWon(Number(c.amount)) : '–'}</td>;
                          })}
                          <td className="text-right py-1.5 pl-2 font-semibold" style={{ color: '#EFC94C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtWon(amount)}</td>
                          <td className="text-center py-1.5 pl-2">
                            <button onClick={() => toggleMemberAllPaid(m.id)} className="rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: paid ? '#12302C' : NEUTRAL_BG, color: paid ? '#7FDCCF' : MUTE }}>{paid ? '완납' : '미납'}</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
        <div className="flex items-center gap-2">
          <span className="text-xs shrink-0" style={{ color: MUTE }}>{nextDinnerRound}차 금액</span>
          <input type="number" value={dinnerAmount} onChange={(e) => setDinnerAmount(e.target.value)} placeholder="금액" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
          <PrimaryBtn onClick={addDinnerRound} icon={Plus}>추가</PrimaryBtn>
        </div>
          </>
        )}
      </Card>

      <Card className="space-y-2">
        <SectionHeader title="지출 내역" open={expenseOpen} onToggle={() => setExpenseOpen((v) => !v)} />
        {expenseOpen && (
          <>
        <div className="flex gap-2">
          <input type="date" value={expDate} onChange={(e) => setExpDate(e.target.value)} className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
          <input value={expDesc} onChange={(e) => setExpDesc(e.target.value)} placeholder="내역 (예: 간식비)" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
        </div>
        <div className="flex gap-2">
          <input type="number" value={expAmount} onChange={(e) => setExpAmount(e.target.value)} placeholder="금액" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
          <PrimaryBtn onClick={addExpense} icon={Plus}>등록</PrimaryBtn>
        </div>
        <div className="pt-2 space-y-1.5" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
          {expensesForMonth.length === 0 && <p className="text-sm py-2" style={{ color: MUTE }}>이번 달 지출 내역이 없어요.</p>}
          {expensesForMonth.map((e) => (
            <div key={e.id} className="flex items-center justify-between text-sm py-1">
              <div className="min-w-0">
                <div style={{ color: INK }}>{e.description}</div>
                <div className="text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(e.date)}{e.recorded_by && ` · ${e.recorded_by}`}</div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <span style={{ color: '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtWon(Number(e.amount))}</span>
                <button onClick={() => requestDelete(() => removeExpense(e.id), '이 지출 내역을 삭제할까요?')} className="p-1" style={{ color: MUTE }}><Trash2 size={14} /></button>
              </div>
            </div>
          ))}
        </div>
          </>
        )}
      </Card>
    </div>
  );
}

function AdminScreen({ members, sessions, checkins, penaltyRule, setPenaltyRule, penaltyCompletions, reload, calendarDays, absenceExcuses, requestDelete, currentMember }) {
  const [date, setDate] = useState(todayStr());
  const session = sessions.find((s) => s.date === date);
  const dayCheckins = session ? checkins.filter((c) => c.session_id === session.id) : [];
  const [manualMemberId, setManualMemberId] = useState(''); const [manualIn, setManualIn] = useState(''); const [manualOut, setManualOut] = useState('');
  const [manualSelectedIds, setManualSelectedIds] = useState([]);
  const [editingRule, setEditingRule] = useState(false); const [ruleInput, setRuleInput] = useState(penaltyRule || '');
  const [expandedPenaltyId, setExpandedPenaltyId] = useState(null);

  // 오늘 접속자수 — 간사만 볼 수 있음, site_visits 테이블에서 오늘 날짜분만 별도 조회(전체 reload 사이클과 무관)
  const isSecretary = currentMember?.role === '간사';
  const [todayVisitCount, setTodayVisitCount] = useState(null);
  useEffect(() => {
    if (!isSecretary) return;
    const start = `${todayStr()}T00:00:00`;
    const end = `${todayStr()}T23:59:59`;
    supabase.from('site_visits').select('id', { count: 'exact', head: true }).gte('visited_at', start).lte('visited_at', end)
      .then(({ count }) => setTodayVisitCount(count ?? 0))
      .catch(() => setTodayVisitCount(null));
  }, [isSecretary]);

  const ensureSession = async () => {
    if (session) return session;
    const s = { id: uid('s'), date, created_at: new Date().toISOString() };
    await insertRow('sessions', s); return s;
  };
  const addManual = async () => {
    if (!manualMemberId || !manualIn) return;
    const s = await ensureSession();
    const inIso = new Date(`${date}T${manualIn}`).toISOString();
    const outIso = manualOut ? new Date(`${date}T${manualOut}`).toISOString() : null;
    await insertRow('checkins', { id: uid('c'), session_id: s.id, member_id: manualMemberId, check_in_at: inIso, check_out_at: outIso });
    await reload();
    setManualMemberId(''); setManualIn(''); setManualOut('');
  };
  const toggleManualSelect = (id) => setManualSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const addManualBulk = async () => {
    if (manualSelectedIds.length === 0 || !manualIn) return;
    const s = await ensureSession();
    const inIso = new Date(`${date}T${manualIn}`).toISOString();
    const outIso = manualOut ? new Date(`${date}T${manualOut}`).toISOString() : null;
    const dayC = checkins.filter((c) => c.session_id === s.id);
    for (const id of manualSelectedIds) {
      const existing = dayC.find((c) => c.member_id === id);
      if (existing) await updateRow('checkins', 'id', existing.id, { check_in_at: inIso, check_out_at: outIso });
      else await insertRow('checkins', { id: uid('c'), session_id: s.id, member_id: id, check_in_at: inIso, check_out_at: outIso });
    }
    await reload();
    setManualSelectedIds([]); setManualIn(''); setManualOut('');
  };
  const updateCheckin = async (id, field, timeVal) => { if (!timeVal) return; await updateRow('checkins', 'id', id, { [field]: new Date(`${date}T${timeVal}`).toISOString() }); await reload(); };
  const removeCheckin = async (id) => { await deleteRow('checkins', 'id', id); await reload(); };
  const saveRule = () => { setPenaltyRule(ruleInput.trim()); setEditingRule(false); };
  const weeklyPenalties = computeWeeklyPenalties(sessions, checkins, calendarDays, members, absenceExcuses);
  const addExcuse = async (memberId, reason) => {
    const existing = absenceExcuses.find((e) => e.date === date && e.member_id === memberId);
    if (existing) await deleteRow('absence_excuses', 'id', existing.id);
    await insertRow('absence_excuses', { id: uid('ae'), date, member_id: memberId, reason });
    await reload();
  };
  const removeExcuse = async (id) => { await deleteRow('absence_excuses', 'id', id); await reload(); };
  const isWeekCompleted = (wk, memberId) => penaltyCompletions.some((p) => p.session_id === wk && p.member_id === memberId);
  const toggleWeekCompletion = async (wk, memberId) => {
    const existing = penaltyCompletions.find((p) => p.session_id === wk && p.member_id === memberId);
    if (existing) await deleteRow('penalty_completions', 'id', existing.id);
    else await insertRow('penalty_completions', { id: uid('p'), session_id: wk, member_id: memberId, completed_at: new Date().toISOString() });
    await reload();
  };
  const weeksWithTargets = weeklyPenalties.filter((w) => w.results.some((r) => r.missedAll));

  const downloadMonthExcel = () => {
    const monthSessions = sessions.filter((s) => s.date.startsWith(date.slice(0, 7)));
    const rows = [];
    monthSessions.forEach((s) => checkins.filter((c) => c.session_id === s.id).forEach((c) => {
      const m = members.find((mm) => mm.id === c.member_id); const dur = durationMin(c.check_in_at, c.check_out_at);
      rows.push({ 날짜: s.date, 이름: m?.name || '', 직급: m?.role || '', 체크인: fmtTime(c.check_in_at), 체크아웃: fmtTime(c.check_out_at), 지속시간_분: dur ?? '', 출석인정: dur !== null && dur >= 30 ? 'O' : 'X' });
    }));
    const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '출석기록'); XLSX.writeFile(wb, `책스초코_출석기록_${date.slice(0, 7)}.xlsx`);
  };

  return (
    <div className="space-y-4">
      {isSecretary && (
        <Card>
          <div className="flex items-center justify-between">
            <span className="text-sm font-semibold" style={{ color: INK }}>오늘 접속자수</span>
            <span className="text-lg font-semibold" style={{ color: '#7FDCCF', fontFamily: "'IBM Plex Mono', monospace" }}>{todayVisitCount === null ? '—' : `${todayVisitCount}회`}</span>
          </div>
          <p className="text-[11px] mt-1" style={{ color: MUTE }}>오늘 앱이 열린 횟수예요 (같은 사람이 여러 번 들어오면 중복 집계될 수 있어요). 간사에게만 보여요.</p>
        </Card>
      )}
      <Card>
        <div className="flex items-center gap-2">
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="rounded-xl border px-3 py-2 text-sm outline-none flex-1" style={inputStyle} />
          <GhostBtn onClick={downloadMonthExcel} icon={Download}>이번 달 다운로드</GhostBtn>
        </div>
      </Card>
      <Card>
        <div className="text-sm font-semibold mb-3" style={{ color: INK }}>{fmtDate(date)} 체크인 목록</div>
        <div className="space-y-2">
          {dayCheckins.length === 0 && <p className="text-sm" style={{ color: MUTE }}>기록이 없어요.</p>}
          {dayCheckins.map((c) => {
            const m = members.find((mm) => mm.id === c.member_id); const dur = durationMin(c.check_in_at, c.check_out_at);
            return (
              <div key={c.id} className="flex flex-wrap items-center gap-2 py-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                <div className="flex items-center gap-2 w-28 shrink-0"><Stamp role={m?.role || '회원'} size={22} tilt={0} /><span className="text-sm truncate" style={{ color: INK }}>{m?.name || '?'}</span></div>
                <input type="time" defaultValue={c.check_in_at ? fmtTime(c.check_in_at) : ''} onBlur={(e) => updateCheckin(c.id, 'check_in_at', e.target.value)} className="rounded-lg border px-2 py-1 text-xs" style={inputStyle} />
                <span className="text-xs" style={{ color: MUTE }}>→</span>
                <input type="time" defaultValue={c.check_out_at ? fmtTime(c.check_out_at) : ''} onBlur={(e) => updateCheckin(c.id, 'check_out_at', e.target.value)} className="rounded-lg border px-2 py-1 text-xs" style={inputStyle} />
                <span className="text-xs font-semibold ml-auto" style={{ color: dur !== null && dur >= 30 ? '#7FDCCF' : '#F0A87C' }}>{dur !== null ? `${dur}분` : '—'}</span>
                <button onClick={() => requestDelete(() => removeCheckin(c.id), '이 출결 기록을 삭제할까요?')} className="p-1" style={{ color: '#F0A87C' }}><Trash2 size={14} /></button>
              </div>
            );
          })}
        </div>
      </Card>
      <Card>
        <div className="text-sm font-semibold mb-1" style={{ color: INK }}>불참 사유 ({fmtDate(date)})</div>
        <p className="text-xs mb-3" style={{ color: MUTE }}>출장·휴가는 그 날 벌칙 판단에서 제외돼요. 개인일정은 사전 파악용으로만 기록되고 결석으로 그대로 집계돼요.</p>
        <div className="space-y-2">
          {members.map((m) => {
            const excuse = absenceExcuses.find((e) => e.date === date && e.member_id === m.id);
            return (
              <div key={m.id} className="flex items-center justify-between py-1.5" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                <div className="flex items-center gap-2 min-w-0"><Stamp role={m.role} size={24} tilt={0} /><span className="text-sm truncate" style={{ color: INK }}>{m.name}</span></div>
                {excuse ? (
                  <div className="flex items-center gap-1.5 shrink-0">
                    <span className="text-xs rounded-full px-2 py-1" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>{excuse.reason}</span>
                    <button onClick={() => removeExcuse(excuse.id)} className="p-1" style={{ color: MUTE }}><X size={14} /></button>
                  </div>
                ) : (
                  <div className="flex gap-1.5 shrink-0 flex-wrap justify-end">
                    <button onClick={() => addExcuse(m.id, '출장')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>출장</button>
                    <button onClick={() => addExcuse(m.id, '휴가')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>휴가</button>
                    <button onClick={() => addExcuse(m.id, '개인일정')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>개인일정</button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </Card>
      <Card className="space-y-2">
        <div className="text-sm font-semibold" style={{ color: INK }}>수동 등록</div>
        <select value={manualMemberId} onChange={(e) => setManualMemberId(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle}>
          <option value="">멤버 선택</option>{members.map((m) => <option key={m.id} value={m.id}>{m.name} ({m.role})</option>)}
        </select>
        <div className="flex gap-2"><input type="time" value={manualIn} onChange={(e) => setManualIn(e.target.value)} className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /><input type="time" value={manualOut} onChange={(e) => setManualOut(e.target.value)} className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
        <PrimaryBtn onClick={addManual} icon={Plus}>등록</PrimaryBtn>

        <div className="pt-3 mt-1" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
          <div className="text-xs mb-2" style={{ color: MUTE }}>일괄 체크 — 여러 명을 한 번에 같은 시간으로 등록해요</div>
          <div className="flex flex-wrap gap-1.5 mb-2">
            {members.map((m) => {
              const checked = manualSelectedIds.includes(m.id);
              return (
                <button key={m.id} onClick={() => toggleManualSelect(m.id)}
                  className="flex items-center gap-1.5 rounded-full border pl-1 pr-2.5 py-1"
                  style={{ borderColor: checked ? '#7FDCCF' : LINE, background: checked ? '#12302C' : 'transparent' }}>
                  <Stamp role={m.role} size={20} tilt={0} />
                  <span className="text-xs" style={{ color: checked ? '#7FDCCF' : INK }}>{m.name}</span>
                </button>
              );
            })}
          </div>
          <PrimaryBtn onClick={addManualBulk} disabled={manualSelectedIds.length === 0 || !manualIn} icon={Check}>{manualSelectedIds.length}명 일괄 등록</PrimaryBtn>
        </div>
      </Card>
      <Card>
        <div className="flex items-center justify-between mb-2">
          <div className="flex items-center gap-1.5 text-sm font-semibold" style={{ color: INK }}><Gavel size={16} style={{ color: '#F0A87C' }} /> 벌칙 관리</div>
          {!editingRule && <button onClick={() => { setRuleInput(penaltyRule || ''); setEditingRule(true); }} className="p-1.5" style={{ color: MUTE }}><Pencil size={14} /></button>}
        </div>
        {editingRule ? (
          <div className="space-y-2 mb-3">
            <textarea value={ruleInput} onChange={(e) => setRuleInput(e.target.value)} rows={2} placeholder="예: 결석 1회당 커피 쏘기" className="w-full rounded-xl border px-3 py-2 text-sm outline-none resize-none" style={inputStyle} />
            <div className="flex gap-2"><PrimaryBtn onClick={saveRule} icon={Check}>저장</PrimaryBtn><GhostBtn onClick={() => setEditingRule(false)} icon={X}>취소</GhostBtn></div>
          </div>
        ) : <p className="text-sm mb-3" style={{ color: penaltyRule ? NEUTRAL_TEXT : MUTE }}>{penaltyRule || '아직 벌칙 규정이 설정되지 않았어요.'}</p>}
        <div className="space-y-3" style={{ borderTop: `1px solid ${ROW_LINE}`, paddingTop: 10 }}>
          {weeksWithTargets.length === 0 && <p className="text-sm" style={{ color: MUTE }}>4일 모두 결석한 벌칙 대상이 없어요.</p>}
          {weeksWithTargets.map((w) => {
            const expanded = expandedPenaltyId === w.weekKey;
            const targets = w.results.filter((r) => r.missedAll);
            const pendingCount = targets.filter((r) => !isWeekCompleted(w.weekKey, r.member.id)).length;
            return (
              <div key={w.weekKey}>
                <button onClick={() => setExpandedPenaltyId(expanded ? null : w.weekKey)} className="w-full flex items-center justify-between py-1.5">
                  <div>
                    <div className="text-sm font-semibold" style={{ color: INK }}>{fmtDate(w.sessions[0].date)} ~ {fmtDate(w.sessions[3].date)}</div>
                    <div className="text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>월~목 4일 모두 독서일</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>대상 {targets.length}명</span>
                    {pendingCount > 0 ? <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}>미이행 {pendingCount}</span> : <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#12302C', color: '#7FDCCF' }}>완료</span>}
                  </div>
                </button>
                {expanded && (
                  <div className="pl-2 pb-2 space-y-1.5">
                    {targets.map((r) => {
                      const done = isWeekCompleted(w.weekKey, r.member.id);
                      return (
                        <div key={r.member.id} className="flex items-center justify-between text-xs py-1">
                          <div className="flex items-center gap-2"><Stamp role={r.member.role} size={22} tilt={0} /><span style={{ color: NEUTRAL_TEXT }}>{r.member.name}</span></div>
                          <button onClick={() => toggleWeekCompletion(w.weekKey, r.member.id)} className="inline-flex items-center gap-1 rounded-full px-2 py-1 font-semibold" style={{ background: done ? '#12302C' : NEUTRAL_BG, color: done ? '#7FDCCF' : MUTE }}>{done ? <Check size={11} /> : null} {done ? '이행 완료' : '미이행'}</button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </Card>
    </div>
  );
}
