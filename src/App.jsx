import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './supabaseClient';
import {
  Crown, Shield, Wallet, User, Plus, Pencil, Trash2, Check, X, Lock, AlertCircle, Mail,
  Megaphone, QrCode, BarChart3, Users, Settings2, Settings, Download, Upload, ChevronLeft, ChevronRight, ChevronUp, ChevronDown, Briefcase, Coffee,
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
// 이 브라우저(기기)를 구분하는 임의의 식별값 — 하드웨어 값이 아니라 우리가 만들어 localStorage에 저장해두는 값
const getDeviceId = () => {
  let id = localStorage.getItem('chexchoco-device-id');
  if (!id) { id = (crypto.randomUUID ? crypto.randomUUID() : uid('dev')); localStorage.setItem('chexchoco-device-id', id); }
  return id;
};
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
const weekKeyOf = (dateStr) => {
  const m = getMonday(dateStr);
  return `${m.getFullYear()}-${pad(m.getMonth() + 1)}-${pad(m.getDate())}`;
};
// 어떤 날짜든 "그 날짜가 속한 달"을 기준으로 몇 월 몇 주차인지 계산 (일요일 시작, 그 달의 첫 완전한 주를 1주차로 삼음)
const monthWeekLabelOf = (dateStr) => {
  const d = new Date(`${dateStr}T00:00:00`);
  const y = d.getFullYear(), mo = d.getMonth();
  const firstOfMonth = new Date(y, mo, 1);
  const firstFullSunday = firstOfMonth.getDay() === 0 ? firstOfMonth : new Date(y, mo, 1 + (7 - firstOfMonth.getDay()));
  const sunday = new Date(d); sunday.setDate(sunday.getDate() - sunday.getDay());
  const week = Math.max(1, Math.floor(Math.round((sunday - firstFullSunday) / 86400000) / 7) + 1);
  return `${mo + 1}월 ${week}주차`;
};
const EXEMPT_EXCUSE_REASONS = ['출장', '휴가']; // 벌칙·출석률 계산에서 제외되는 사유 (업무·개인일정은 제외 안 됨 — 결석으로 그대로 집계)

// 주 단위 벌칙 계산: 월~목 4일이 모두 독서일이고 이미 다 지난 "완결된 주"에서,
// 4일간 출석 환산 합계가 1일 미만(= 30분 이상 출석이 하나도 없고, 15~29분 출석도 2회 미만)인 멤버만 그 주의 벌칙 대상이 됨.
// (15~29분 출석은 0.5일로 환산되므로, 그런 날이 2번이면 1일로 합산되어 벌칙에서 제외됨)
const computeWeeklyPenalties = (sessions, checkins, calendarDays, members, absenceExcuses = []) => {
  const now = new Date();
  // 세션 레코드가 우연히 생성 안 된 날이 있어도 결석으로 정확히 판정되도록, calendarDays에 지정된 '독서일' 날짜를 기준으로 주를 구성
  const readingDates = calendarDays.filter((d) => d.type === '독서일').map((d) => d.date);
  const byWeek = {};
  readingDates.forEach((ds) => {
    if (!isMonToThu(ds)) return;
    const wk = weekKeyOf(ds);
    (byWeek[wk] = byWeek[wk] || new Set()).add(ds);
  });
  return Object.entries(byWeek)
    .filter(([wk, dateSet]) => {
      if (dateSet.size !== 4) return false; // 월~목 4일이 전부 독서일로 지정된 "정상 주"만
      if (!weekQualifiesForPenalty(wk, calendarDays)) return false;
      const thuCutoff = new Date(`${wk}T00:00:00`); thuCutoff.setDate(thuCutoff.getDate() + 3); thuCutoff.setHours(13, 0, 0, 0); // 그 주 목요일 13:00
      return now >= thuCutoff; // 목요일 13시 이후에만 벌칙 확정 (그 전엔 진행 중인 주로 보류)
    })
    .map(([wk, dateSet]) => {
      const datesSorted = [...dateSet].sort();
      const results = members.map((m) => {
        // 출장/휴가/업무 사유가 있는 날만 그 멤버에 한해 판단 대상에서 제외 (개인일정은 제외 안 됨)
        const relevant = datesSorted.filter((ds) => !absenceExcuses.some((e) => e.date === ds && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason)));
        if (relevant.length === 0) return { member: m, missedAll: false }; // 4일 다 사유 있으면 벌칙 대상 아님
        const totalEquivalent = relevant.reduce((sum, ds) => {
          const s = sessions.find((ss) => ss.date === ds); // 그 날 세션 레코드가 아예 없으면 = 아무도 체크인 안 한 것 = 결석 처리
          const c = s ? checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id) : null;
          const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
          return sum + attendanceEquivalent(dur);
        }, 0);
        return { member: m, missedAll: totalEquivalent < 1 };
      });
      return { weekKey: wk, sessions: datesSorted, results };
    })
    .sort((a, b) => b.weekKey.localeCompare(a.weekKey));
};

/* ---------- Supabase data layer ---------- */
const TABLES = ['members', 'notices', 'notice_views', 'sessions', 'checkins', 'penalty_completions', 'calendar_days', 'settings', 'photos', 'absence_excuses', 'meeting_locations', 'dues_payments', 'expenses', 'dinner_collections', 'book_shares', 'notifications', 'book_tower_entries'];

async function fetchAll(tables = TABLES) {
  // members는 pin 컬럼이 빠진 members_public 뷰에서 조회 (일반 조회 시 PIN이 클라이언트로 전송되지 않도록)
  const results = await Promise.all(tables.map((t) => supabase.from(t === 'members' ? 'members_public' : t).select('*')));
  const out = {};
  tables.forEach((t, i) => { out[t] = results[i].data || []; });
  return out;
}
// PIN 검증 전용 — 필요한 순간에만, 그 멤버 한 명의 pin만 좁게 조회해서 비교 (평소 목록 조회엔 PIN이 포함되지 않음)
async function verifyPin(memberId, input) {
  const { data } = await supabase.from('members').select('pin').eq('id', memberId).maybeSingle();
  const actual = data?.pin || null;
  if (!actual) return true; // PIN 미설정 멤버는 항상 통과
  return input === actual;
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
  const [bookShares, setBookShares] = useState([]);
  const [notifications, setNotifications] = useState([]);
  const [bookTowerEntries, setBookTowerEntries] = useState([]);
  const [showNotifications, setShowNotifications] = useState(false);
  const [currentUserId, setCurrentUserId] = useState(() => localStorage.getItem('chexchoco-current-user') || null);
  // 기기 등록: null=조회중, false=미등록, 문자열=등록된 멤버 id (한번 정해지면 앱에서는 절대 못 바꿈 — DB에 update/delete 정책 자체가 없음)
  const [deviceRegMemberId, setDeviceRegMemberId] = useState(null);
  const [registerOfferMember, setRegisterOfferMember] = useState(null); // 방금 로그인한 멤버 — "이 기기 등록할까요?" 제안용
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

  const reload = async (tables) => {
    try {
      const data = await fetchAll(tables);
      if (data.members) setMembers(data.members);
      if (data.notices) setNotices(data.notices);
      if (data.notice_views) setNoticeViews(data.notice_views);
      if (data.sessions) setSessions(data.sessions);
      if (data.checkins) setCheckins(data.checkins);
      if (data.penalty_completions) setPenaltyCompletions(data.penalty_completions);
      if (data.calendar_days) setCalendarDays(data.calendar_days);
      if (data.settings) setSettings(data.settings);
      if (data.photos) setPhotos(data.photos);
      if (data.absence_excuses) setAbsenceExcuses(data.absence_excuses);
      if (data.meeting_locations) setMeetingLocations(data.meeting_locations);
      if (data.dues_payments) setDuesPayments(data.dues_payments);
      if (data.expenses) setExpenses(data.expenses);
      if (data.dinner_collections) setDinnerCollections(data.dinner_collections);
      if (data.book_shares) setBookShares(data.book_shares);
      if (data.notifications) setNotifications(data.notifications);
      if (data.book_tower_entries) setBookTowerEntries(data.book_tower_entries);
      setError('');
    } catch (e) { setError('데이터를 불러오지 못했어요. 새로고침해 주세요.'); }
    setLoaded(true);
  };
  useEffect(() => { reload(); }, []);
  // 기기 등록 여부 확인 — 등록돼 있으면 로그인 절차 없이 자동으로 그 멤버로 인식
  useEffect(() => {
    const deviceId = getDeviceId();
    supabase.from('device_registrations').select('member_id').eq('device_id', deviceId).maybeSingle()
      .then(({ data }) => {
        if (data?.member_id) { setDeviceRegMemberId(data.member_id); setIdentity(data.member_id); }
        else setDeviceRegMemberId(false);
      })
      .catch(() => setDeviceRegMemberId(false));
  }, []);
  // 오늘 접속자수 집계용 — 페이지 로드마다 방문 기록 1건 남김 (site_visits 테이블, 전체 reload 사이클과는 무관하게 별도 처리)
  // 로그인 상태라면 어떤 멤버인지도 같이 기록 (localStorage에서 즉시 읽히므로 currentUserId는 마운트 시점에 이미 확정됨)
  useEffect(() => {
    supabase.from('site_visits').insert({ id: uid('visit'), visited_at: new Date().toISOString(), member_id: currentUserId || null }).then(({ error }) => { if (error) console.error('site_visits insert failed:', error); }).catch((e) => console.error('site_visits insert failed:', e));
    // 30일 지난 방문 기록은 자동 정리 (테이블이 무기한 쌓이지 않도록, 페이지 로드마다 가볍게 체크)
    const cutoff = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    supabase.from('site_visits').delete().lt('visited_at', cutoff).then(() => {}).catch(() => {});
  }, []);

  const penaltyRule = settings.find((s) => s.key === 'penaltyRule')?.value || '';
  const setPenaltyRule = async (v) => { await upsertRow('settings', { key: 'penaltyRule', value: v }, 'key'); reload(); };

  const setIdentity = (id) => {
    if (id) localStorage.setItem('chexchoco-current-user', id); else localStorage.removeItem('chexchoco-current-user');
    setCurrentUserId(id);
  };
  const openLogin = () => { setShowLoginModal(true); setModalSelectedId(''); setModalPinInput(''); setModalPinError(''); };
  const closeLogin = () => { setShowLoginModal(false); setModalSelectedId(''); setModalPinInput(''); setModalPinError(''); };
  const logout = () => { setIdentity(null); };
  const submitLogin = async () => {
    const m = members.find((mm) => mm.id === modalSelectedId);
    if (!m) { setModalPinError('사용자를 선택해 주세요.'); return; }
    const ok = await verifyPin(m.id, modalPinInput);
    if (!ok) { setModalPinError('PIN이 일치하지 않아요.'); return; }
    setIdentity(m.id);
    closeLogin();
    if (deviceRegMemberId === false) setRegisterOfferMember(m); // 이 기기가 아직 미등록이면 등록 제안
  };
  const registerThisDevice = async () => {
    if (!registerOfferMember) return;
    const deviceId = getDeviceId();
    await supabase.from('device_registrations').insert({ device_id: deviceId, member_id: registerOfferMember.id });
    setDeviceRegMemberId(registerOfferMember.id);
    setRegisterOfferMember(null);
  };

  const currentMember = members.find((m) => m.id === currentUserId) || null;
  const myNotifications = currentMember ? notifications.filter((n) => n.member_id === currentMember.id).sort((a, b) => b.created_at.localeCompare(a.created_at)) : [];
  const openNotification = async (n) => {
    if (!n.read_at) await updateRow('notifications', 'id', n.id, { read_at: new Date().toISOString() });
    setShowNotifications(false);
    setTab('gallery');
    await reload();
  };
  const markAllNotificationsRead = async () => {
    const unread = myNotifications.filter((n) => !n.read_at);
    for (const n of unread) await updateRow('notifications', 'id', n.id, { read_at: new Date().toISOString() });
    if (unread.length) await reload();
  };

  // 삭제 시 실수 방지용 재확인(로그인 PIN) — PIN이 설정된 계정이면 PIN 입력, 아니면 한 번 더 확인만
  const requestDelete = (onConfirm, message = '정말 삭제할까요?') => { setPendingDelete({ onConfirm, message }); setDeletePinInput(''); setDeletePinError(''); };
  const cancelDelete = () => { setPendingDelete(null); setDeletePinInput(''); setDeletePinError(''); };
  const confirmDelete = async () => {
    if (currentMember?.has_pin) {
      const ok = await verifyPin(currentMember.id, deletePinInput);
      if (!ok) { setDeletePinError('PIN이 일치하지 않아요.'); return; }
    }
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
    { key: 'dashboard', label: '현황', icon: BarChart3 },
    { key: 'gallery', label: '서재', icon: ImageIcon },
    { key: 'users', label: '멤버', icon: Users },
  ];

  if (!loaded) {
    return <div className="min-h-screen flex items-center justify-center" style={{ background: PAPER_BG }}>
      <div className="text-sm" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>불러오는 중…</div>
    </div>;
  }

  return (
    <div className="min-h-screen" style={{ background: PAPER_BG, fontFamily: "'Inter', sans-serif" }}>
      <style>{`
        input:focus-visible, select:focus-visible, textarea:focus-visible,
        button:focus-visible, a:focus-visible, [tabindex]:focus-visible {
          outline: 2px solid #7FA8D9;
          outline-offset: 2px;
        }
      `}</style>
      <div className="max-w-3xl mx-auto px-4 pt-6 pb-24">
        <div className="mb-6">
          <div className="flex items-center justify-between mb-1.5">
            <span className="text-[11px] tracking-[0.2em] uppercase" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>KEPCO Reading Club</span>
            <div className="flex items-center gap-1.5">
              {canManageUsers && (
                <button onClick={() => setTab('treasury')} aria-label="회계"
                  className="p-1.5 rounded-full" style={{ background: tab === 'treasury' ? BTN_BG : NEUTRAL_BG, color: tab === 'treasury' ? BTN_TEXT : MUTE }}>
                  <Wallet size={15} />
                </button>
              )}
              {currentMember && (
                <button onClick={() => setShowNotifications(true)} aria-label="알림함" className="relative p-1.5 rounded-full" style={{ background: NEUTRAL_BG, color: MUTE }}>
                  <Mail size={15} />
                  {notifications.filter((n) => n.member_id === currentMember.id && !n.read_at).length > 0 && (
                    <span className="absolute -top-0.5 -right-0.5 rounded-full flex items-center justify-center text-[9px] font-bold" style={{ width: 14, height: 14, background: '#E5484D', color: '#fff' }}>
                      {notifications.filter((n) => n.member_id === currentMember.id && !n.read_at).length}
                    </span>
                  )}
                </button>
              )}
              {canManageAttendance && (
                <button onClick={() => setTab('admin')} aria-label="설정"
                  className="p-1.5 rounded-full" style={{ background: tab === 'admin' ? BTN_BG : NEUTRAL_BG, color: tab === 'admin' ? BTN_TEXT : MUTE }}>
                  <Settings size={15} />
                </button>
              )}
              {deviceRegMemberId ? (
                <span className="flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold" style={{ background: BTN_BG, color: BTN_TEXT }} title="이 기기는 등록된 사용자 전용이에요. 다른 사람으로 전환할 수 없어요.">
                  <Lock size={11} />{currentMember ? <><Stamp role={currentMember.role} size={16} tilt={0} />{currentMember.name}님</> : '등록된 기기'}
                </span>
              ) : (
                <button onClick={() => (currentMember ? logout() : openLogin())}
                  className="flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold"
                  style={{ background: currentMember ? BTN_BG : NEUTRAL_BG, color: currentMember ? BTN_TEXT : NEUTRAL_TEXT }}>
                  {currentMember ? <><Stamp role={currentMember.role} size={16} tilt={0} />{currentMember.name}님 · 로그아웃</> : <>로그인</>}
                </button>
              )}
            </div>
          </div>
          <h1 className="text-center text-4xl font-semibold mt-3" style={{ fontFamily: "'Fraunces', serif", color: INK }}>책스초코</h1>
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
                <button onClick={closeLogin} aria-label="닫기"><X size={18} style={{ color: MUTE }} /></button>
              </div>
              {sortedMembers.length === 0 ? (
                <p className="text-sm" style={{ color: MUTE }}>등록된 멤버가 없어요. 사용자관리에서 첫 멤버를 등록해 주세요.</p>
              ) : (
                <div className="space-y-3">
                  <div>
                    <label htmlFor="login-member-select" className="text-xs mb-1 block" style={{ color: MUTE }}>이름 선택</label>
                    <select id="login-member-select" value={modalSelectedId} onChange={(e) => { setModalSelectedId(e.target.value); setModalPinInput(''); setModalPinError(''); }}
                      className="w-full rounded-xl border px-3 py-2.5 text-sm outline-none" style={inputStyle}>
                      <option value="">— 선택하세요 —</option>
                      {sortedMembers.map((m) => <option key={m.id} value={m.id}>{m.name} ({m.role})</option>)}
                    </select>
                  </div>
                  {modalSelectedId && (() => {
                    const m = sortedMembers.find((mm) => mm.id === modalSelectedId);
                    return m?.has_pin ? (
                      <div>
                        <label htmlFor="login-pin-input" className="text-xs mb-1 block" style={{ color: MUTE }}>PIN</label>
                        <input id="login-pin-input" type="password" inputMode="numeric" maxLength={4} value={modalPinInput}
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

        {registerOfferMember && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.7)' }}>
            <div className="w-full max-w-sm rounded-2xl border p-5" style={{ background: CARD_BG, borderColor: LINE }}>
              <div className="flex items-center gap-1.5 text-sm font-semibold mb-2" style={{ color: INK }}><Lock size={15} /> 이 기기를 등록할까요?</div>
              <p className="text-sm mb-2" style={{ color: NEUTRAL_TEXT }}>이 기기를 <strong>{registerOfferMember.name}님 전용</strong>으로 등록하면, 다음부터 로그인 없이 자동으로 인식돼요.</p>
              <p className="text-xs mb-4" style={{ color: '#F0A87C' }}>⚠️ 한 번 등록하면 앱에서는 되돌릴 수 없어요. 다른 사람이 이 기기로 로그인할 수 없게 돼요.</p>
              <div className="flex gap-2">
                <PrimaryBtn onClick={registerThisDevice} icon={Lock}>등록하기</PrimaryBtn>
                <GhostBtn onClick={() => setRegisterOfferMember(null)}>나중에</GhostBtn>
              </div>
            </div>
          </div>
        )}

        {showNotifications && currentMember && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.7)' }} onClick={() => setShowNotifications(false)}>
            <div className="w-full max-w-sm rounded-2xl border p-4" style={{ background: CARD_BG, borderColor: LINE, maxHeight: '75vh', overflowY: 'auto' }} onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between mb-3">
                <div className="flex items-center gap-1.5 text-sm font-semibold" style={{ color: INK }}><Mail size={15} /> 알림함</div>
                <div className="flex items-center gap-2">
                  {myNotifications.some((n) => !n.read_at) && (
                    <button onClick={markAllNotificationsRead} className="text-[11px] underline underline-offset-2" style={{ color: MUTE }}>모두 읽음</button>
                  )}
                  <button onClick={() => setShowNotifications(false)} aria-label="닫기"><X size={16} style={{ color: MUTE }} /></button>
                </div>
              </div>
              {myNotifications.length === 0 ? (
                <p className="text-sm text-center py-6" style={{ color: MUTE }}>알림이 없어요.</p>
              ) : (
                <div className="space-y-1.5">
                  {myNotifications.map((n) => (
                    <button key={n.id} onClick={() => openNotification(n)} className="w-full text-left rounded-xl p-2.5" style={{ background: n.read_at ? 'transparent' : 'rgba(240,168,124,0.08)', border: `1px solid ${n.read_at ? LINE : 'rgba(240,168,124,0.3)'}` }}>
                      <div className="text-xs" style={{ color: n.read_at ? MUTE : INK }}>{n.message}</div>
                      <div className="text-[10px] mt-0.5" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(n.created_at.slice(0, 10))}</div>
                    </button>
                  ))}
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
                <button onClick={cancelDelete} aria-label="닫기"><X size={18} style={{ color: MUTE }} /></button>
              </div>
              <p className="text-sm mb-3" style={{ color: NEUTRAL_TEXT }}>{pendingDelete.message}</p>
              {currentMember?.has_pin ? (
                <div className="mb-3">
                  <label htmlFor="delete-pin-input" className="text-xs mb-1 block" style={{ color: MUTE }}>로그인 PIN 확인</label>
                  <input id="delete-pin-input" type="password" inputMode="numeric" maxLength={4} value={deletePinInput}
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
        {tab === 'gallery' && <GalleryScreen photos={photos} currentMember={currentMember} canManage={canManageUsers} reload={reload} members={members} sessions={sessions} checkins={checkins} requestDelete={requestDelete} showToast={showToast} bookShares={bookShares} bookTowerEntries={bookTowerEntries} />}
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
  const isSecretary = currentMember?.role === '간사'; // 공지 조회수는 간사에게만 노출

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
            <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="제목" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="공지 제목" />
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
                  <button onClick={() => moveNotice(n.id, 'up')} className="p-1.5" style={{ color: MUTE }} aria-label="공지 위로 이동"><ChevronUp size={15} /></button>
                  <button onClick={() => moveNotice(n.id, 'down')} className="p-1.5" style={{ color: MUTE }} aria-label="공지 아래로 이동"><ChevronDown size={15} /></button>
                  <button onClick={() => startEdit(n)} className="p-1.5" style={{ color: MUTE }} aria-label="공지 수정"><Pencil size={15} /></button>
                  <button onClick={() => requestDelete(() => remove(n.id), '이 공지사항을 삭제할까요?')} className="p-1.5" style={{ color: '#F0A87C' }} aria-label="공지 삭제"><Trash2 size={15} /></button>
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
                  {isSecretary && (
                    <button onClick={() => setExpandedViewsId(expanded ? null : n.id)} className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}>
                      <Eye size={12} /> {views.length}명 조회{views.length > 0 ? (expanded ? ' 숨기기' : ' 보기') : ''}
                    </button>
                  )}
                </div>
                {isSecretary && expanded && (
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
function GalleryScreen({ photos, currentMember, canManage, reload, members, sessions, checkins, requestDelete, showToast, bookShares, bookTowerEntries }) {
  const [error, setError] = useState('');
  const [uploading, setUploading] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const [editingDate, setEditingDate] = useState(false);
  const [dateInput, setDateInput] = useState('');
  const [editingCaption, setEditingCaption] = useState(false);
  const [captionInput, setCaptionInput] = useState('');
  const [photoPage, setPhotoPage] = useState(0);
  const PHOTOS_PER_PAGE = 9;
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
  const totalPhotoPages = Math.max(1, Math.ceil(filtered.length / PHOTOS_PER_PAGE));
  const safePage = Math.min(photoPage, totalPhotoPages - 1);
  const pagedPhotos = filtered.slice(safePage * PHOTOS_PER_PAGE, safePage * PHOTOS_PER_PAGE + PHOTOS_PER_PAGE);

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

  // ---------- 도서 공유함 ----------
  const [showShareForm, setShowShareForm] = useState(false);
  const [shareKind, setShareKind] = useState('offer');
  const [shareTitle, setShareTitle] = useState('');
  const [shareAuthor, setShareAuthor] = useState('');
  const [sharePublisher, setSharePublisher] = useState('');
  const [viewingShareId, setViewingShareId] = useState(null);
  const [dueDateInput, setDueDateInput] = useState('');
  const notify = async (memberId, message, linkId) => {
    await insertRow('notifications', { id: uid('nt'), member_id: memberId, message, link_id: linkId, created_at: new Date().toISOString() });
  };
  const submitBookShare = async () => {
    if (!currentMember || !shareTitle.trim()) return;
    await insertRow('book_shares', {
      id: uid('bs'), kind: shareKind, posted_by: currentMember.id,
      book_title: shareTitle.trim(), book_author: shareAuthor.trim() || null, book_publisher: sharePublisher.trim() || null,
      status: 'open', created_at: new Date().toISOString(),
    });
    setShareTitle(''); setShareAuthor(''); setSharePublisher(''); setShowShareForm(false);
    await reload();
  };
  const respondToShare = async (share) => {
    if (!currentMember) return;
    const today = todayStr();
    const due = new Date(); due.setDate(due.getDate() + 14);
    const dueStr = `${due.getFullYear()}-${pad(due.getMonth() + 1)}-${pad(due.getDate())}`;
    await updateRow('book_shares', 'id', share.id, { status: 'matched', matched_by: currentMember.id, borrowed_at: today, due_date: dueStr });
    await notify(share.posted_by, `${currentMember.name}님이 [${share.book_title}] ${share.kind === 'offer' ? '제공' : '요청'}에 응답했어요.`, share.id);
    await reload();
  };
  const updateDueDate = async (share, newDate) => {
    await updateRow('book_shares', 'id', share.id, { due_date: newDate });
    await reload();
  };
  const randomPastel = () => {
    const hues = [200, 150, 30, 340, 260, 100, 20, 280, 180];
    const h = hues[Math.floor(Math.random() * hues.length)];
    return `hsl(${h}, 55%, 68%)`;
  };
  const confirmReturn = async (share) => {
    const borrowerId = share.kind === 'offer' ? share.matched_by : share.posted_by;
    const today = todayStr();
    await updateRow('book_shares', 'id', share.id, { status: 'returned', returned_at: today });
    await insertRow('book_tower_entries', { id: uid('bt'), member_id: borrowerId, book_title: share.book_title, finished_date: today, color: randomPastel(), source_share_id: share.id, created_at: new Date().toISOString() });
    await notify(borrowerId, `[${share.book_title}] 반납 완료 처리됐어요. 내 책탑에 추가됐어요.`, share.id);
    await reload();
  };
  const deleteBookShare = async (share) => {
    await deleteRow('book_shares', 'id', share.id);
    setViewingShareId(null);
    await reload();
  };
  const offers = bookShares.filter((s) => s.kind === 'offer').sort((a, b) => b.created_at.localeCompare(a.created_at));
  const requests = bookShares.filter((s) => s.kind === 'request').sort((a, b) => b.created_at.localeCompare(a.created_at));
  const viewingShare = bookShares.find((s) => s.id === viewingShareId) || null;

  // ---------- 책탑 ----------
  const [towerView, setTowerView] = useState('mine'); // 'mine' | 'group'
  const [showTowerAdd, setShowTowerAdd] = useState(false);
  const [towerTitleInput, setTowerTitleInput] = useState('');
  const [towerDateInput, setTowerDateInput] = useState(todayStr());
  const addManualTowerEntry = async () => {
    if (!currentMember || !towerTitleInput.trim()) return;
    await insertRow('book_tower_entries', { id: uid('bt'), member_id: currentMember.id, book_title: towerTitleInput.trim(), finished_date: towerDateInput || todayStr(), color: randomPastel(), source_share_id: null, created_at: new Date().toISOString() });
    setTowerTitleInput(''); setTowerDateInput(todayStr()); setShowTowerAdd(false);
    await reload();
  };
  const removeTowerEntry = async (entryId) => { await deleteRow('book_tower_entries', 'id', entryId); await reload(); };
  const myTower = currentMember ? bookTowerEntries.filter((t) => t.member_id === currentMember.id).sort((a, b) => a.finished_date.localeCompare(b.finished_date)) : [];
  const groupTower = [...bookTowerEntries].sort((a, b) => a.finished_date.localeCompare(b.finished_date));

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
          {pagedPhotos.map((p) => (
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
      {filtered.length > PHOTOS_PER_PAGE && (
        <div className="flex items-center justify-center gap-3">
          <button onClick={() => setPhotoPage((p) => Math.max(0, p - 1))} disabled={safePage === 0} aria-label="이전 페이지" className="p-1.5 rounded-full" style={{ color: safePage === 0 ? LINE : MUTE }}><ChevronLeft size={16} /></button>
          <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{safePage + 1} / {totalPhotoPages}</span>
          <button onClick={() => setPhotoPage((p) => Math.min(totalPhotoPages - 1, p + 1))} disabled={safePage >= totalPhotoPages - 1} aria-label="다음 페이지" className="p-1.5 rounded-full" style={{ color: safePage >= totalPhotoPages - 1 ? LINE : MUTE }}><ChevronRight size={16} /></button>
        </div>
      )}

      {/* ---------- 도서 공유함 ---------- */}
      <Card>
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-1.5 text-sm font-semibold" style={{ color: INK }}><BookOpen size={16} style={{ color: '#7FA8D9' }} /> 도서 공유함</div>
          {currentMember && <button onClick={() => setShowShareForm((v) => !v)} className="text-xs rounded-full px-3 py-1.5 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>{showShareForm ? '취소' : '글쓰기'}</button>}
        </div>
        {showShareForm && (
          <div className="space-y-2 mb-3 pb-3" style={{ borderBottom: `1px solid ${ROW_LINE}` }}>
            <div className="flex gap-2">
              <button onClick={() => setShareKind('offer')} className="flex-1 rounded-xl py-2 text-xs font-semibold" style={{ background: shareKind === 'offer' ? BTN_BG : NEUTRAL_BG, color: shareKind === 'offer' ? BTN_TEXT : NEUTRAL_TEXT }}>책 빌려줄까요?</button>
              <button onClick={() => setShareKind('request')} className="flex-1 rounded-xl py-2 text-xs font-semibold" style={{ background: shareKind === 'request' ? BTN_BG : NEUTRAL_BG, color: shareKind === 'request' ? BTN_TEXT : NEUTRAL_TEXT }}>책 빌려주실 수 있나요?</button>
            </div>
            <input value={shareTitle} onChange={(e) => setShareTitle(e.target.value)} placeholder="책 제목 (필수)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <div className="grid grid-cols-2 gap-2">
              <input value={shareAuthor} onChange={(e) => setShareAuthor(e.target.value)} placeholder="저자 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
              <input value={sharePublisher} onChange={(e) => setSharePublisher(e.target.value)} placeholder="출판사 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            </div>
            <PrimaryBtn onClick={submitBookShare} icon={Plus}>등록</PrimaryBtn>
          </div>
        )}
        {!currentMember && !showShareForm && <p className="text-xs mb-2" style={{ color: MUTE }}>상단에서 본인을 먼저 선택해야 글을 올릴 수 있어요.</p>}
        <div className="mb-1.5 text-xs font-semibold" style={{ color: MUTE }}>📚 책 빌려줄까요? ({offers.length})</div>
        <div className="space-y-1.5 mb-3">
          {offers.length === 0 && <p className="text-xs" style={{ color: MUTE }}>등록된 글이 없어요.</p>}
          {offers.map((s) => {
            const poster = members.find((m) => m.id === s.posted_by);
            return (
              <button key={s.id} onClick={() => { setViewingShareId(s.id); setDueDateInput(''); }} className="w-full flex items-center justify-between rounded-xl px-3 py-2 text-left" style={{ background: NEUTRAL_BG }}>
                <div className="min-w-0">
                  <div className="text-sm truncate" style={{ color: INK }}>{s.book_title}</div>
                  <div className="text-[10px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{dispName(poster?.name || '', isLoggedIn)}</div>
                </div>
                <span className="text-[10px] rounded-full px-2 py-0.5 font-semibold shrink-0 ml-2" style={
                  s.status === 'open' ? { background: '#12302C', color: '#7FDCCF' } : s.status === 'matched' ? { background: '#3A2E10', color: '#EFC94C' } : { background: NEUTRAL_BG, color: MUTE, border: `1px solid ${LINE}` }
                }>{s.status === 'open' ? '대여가능' : s.status === 'matched' ? '대여중' : '반납완료'}</span>
              </button>
            );
          })}
        </div>
        <div className="mb-1.5 text-xs font-semibold" style={{ color: MUTE }}>🙋 책 빌려주실 수 있나요? ({requests.length})</div>
        <div className="space-y-1.5">
          {requests.length === 0 && <p className="text-xs" style={{ color: MUTE }}>등록된 글이 없어요.</p>}
          {requests.map((s) => {
            const poster = members.find((m) => m.id === s.posted_by);
            return (
              <button key={s.id} onClick={() => { setViewingShareId(s.id); setDueDateInput(''); }} className="w-full flex items-center justify-between rounded-xl px-3 py-2 text-left" style={{ background: NEUTRAL_BG }}>
                <div className="min-w-0">
                  <div className="text-sm truncate" style={{ color: INK }}>{s.book_title}</div>
                  <div className="text-[10px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{dispName(poster?.name || '', isLoggedIn)}</div>
                </div>
                <span className="text-[10px] rounded-full px-2 py-0.5 font-semibold shrink-0 ml-2" style={
                  s.status === 'open' ? { background: '#12302C', color: '#7FDCCF' } : s.status === 'matched' ? { background: '#3A2E10', color: '#EFC94C' } : { background: NEUTRAL_BG, color: MUTE, border: `1px solid ${LINE}` }
                }>{s.status === 'open' ? '요청중' : s.status === 'matched' ? '대여중' : '반납완료'}</span>
              </button>
            );
          })}
        </div>
      </Card>

      {/* ---------- 도서 공유함 상세보기 ---------- */}
      {viewingShare && (() => {
        const poster = members.find((m) => m.id === viewingShare.posted_by);
        const matcher = members.find((m) => m.id === viewingShare.matched_by);
        const ownerId = viewingShare.kind === 'offer' ? viewingShare.posted_by : viewingShare.matched_by;
        const borrowerId = viewingShare.kind === 'offer' ? viewingShare.matched_by : viewingShare.posted_by;
        const isOwner = currentMember?.id === ownerId;
        return (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.7)' }} onClick={() => setViewingShareId(null)}>
            <div className="w-full max-w-sm rounded-2xl border p-5" style={{ background: CARD_BG, borderColor: LINE }} onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between mb-3">
                <span className="text-[11px] rounded-full px-2 py-0.5 font-semibold" style={{ background: viewingShare.kind === 'offer' ? '#1E2A38' : '#332815', color: viewingShare.kind === 'offer' ? '#7FA8D9' : '#EFC94C' }}>{viewingShare.kind === 'offer' ? '제공' : '요청'}</span>
                <button onClick={() => setViewingShareId(null)} aria-label="닫기"><X size={16} style={{ color: MUTE }} /></button>
              </div>
              <div className="text-base font-semibold mb-2" style={{ color: INK }}>{viewingShare.book_title}</div>
              <div className="text-xs space-y-1 mb-3" style={{ color: NEUTRAL_TEXT }}>
                {viewingShare.book_author && <div>저자: {viewingShare.book_author}</div>}
                {viewingShare.book_publisher && <div>출판사: {viewingShare.book_publisher}</div>}
                <div style={{ color: MUTE }}>글쓴이: {dispName(poster?.name || '', isLoggedIn)}</div>
              </div>
              {viewingShare.status === 'open' && (
                <>
                  {currentMember && currentMember.id !== viewingShare.posted_by ? (
                    <PrimaryBtn onClick={() => { respondToShare(viewingShare); setViewingShareId(null); }} icon={Check}>{viewingShare.kind === 'offer' ? '제가 빌릴게요' : '제가 빌려드릴게요'}</PrimaryBtn>
                  ) : currentMember?.id === viewingShare.posted_by ? (
                    <p className="text-xs" style={{ color: MUTE }}>다른 회원의 응답을 기다리는 중이에요.</p>
                  ) : null}
                </>
              )}
              {viewingShare.status !== 'open' && (
                <div className="space-y-2 pt-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="text-xs" style={{ color: NEUTRAL_TEXT }}>응답: {dispName(matcher?.name || '', isLoggedIn)}</div>
                  <div className="text-xs" style={{ color: MUTE }}>대여일: {fmtDate(viewingShare.borrowed_at)}</div>
                  {viewingShare.status === 'matched' && (
                    <div className="flex items-center gap-1.5">
                      <span className="text-xs" style={{ color: MUTE }}>반납기한:</span>
                      {isOwner || currentMember?.id === borrowerId ? (
                        <input type="date" value={dueDateInput || viewingShare.due_date || ''} onChange={(e) => setDueDateInput(e.target.value)} onBlur={() => dueDateInput && updateDueDate(viewingShare, dueDateInput)}
                          className="rounded-lg border px-1.5 py-1 text-[11px] outline-none" style={inputStyle} aria-label="반납기한" />
                      ) : (
                        <span className="text-xs" style={{ color: NEUTRAL_TEXT }}>{fmtDate(viewingShare.due_date)}</span>
                      )}
                    </div>
                  )}
                  {viewingShare.status === 'returned' && <div className="text-xs" style={{ color: '#7FDCCF' }}>✓ 반납완료 · {fmtDate(viewingShare.returned_at)}</div>}
                  {viewingShare.status === 'matched' && isOwner && (
                    <PrimaryBtn onClick={() => { confirmReturn(viewingShare); setViewingShareId(null); }} icon={Check}>반납 완료 처리</PrimaryBtn>
                  )}
                </div>
              )}
              {(currentMember?.id === viewingShare.posted_by || canManage) && viewingShare.status === 'open' && (
                <button onClick={() => deleteBookShare(viewingShare)} className="text-[11px] underline underline-offset-2 mt-3" style={{ color: MUTE }}>글 삭제</button>
              )}
            </div>
          </div>
        );
      })()}

      {/* ---------- 책탑 ---------- */}
      <Card>
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-1.5 text-sm font-semibold" style={{ color: INK }}>📚 책탑</div>
          {currentMember && <button onClick={() => setShowTowerAdd((v) => !v)} className="text-xs rounded-full px-3 py-1.5 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>{showTowerAdd ? '취소' : '+ 책 추가'}</button>}
        </div>
        {showTowerAdd && (
          <div className="space-y-2 mb-3 pb-3" style={{ borderBottom: `1px solid ${ROW_LINE}` }}>
            <input value={towerTitleInput} onChange={(e) => setTowerTitleInput(e.target.value)} placeholder="다 읽은 책 제목" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <input type="date" value={towerDateInput} onChange={(e) => setTowerDateInput(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <PrimaryBtn onClick={addManualTowerEntry} icon={Plus}>추가</PrimaryBtn>
          </div>
        )}
        <div className="flex gap-2 mb-3">
          <button onClick={() => setTowerView('mine')} className="flex-1 rounded-xl py-1.5 text-xs font-semibold" style={{ background: towerView === 'mine' ? BTN_BG : NEUTRAL_BG, color: towerView === 'mine' ? BTN_TEXT : NEUTRAL_TEXT }}>내 책탑 ({myTower.length})</button>
          <button onClick={() => setTowerView('group')} className="flex-1 rounded-xl py-1.5 text-xs font-semibold" style={{ background: towerView === 'group' ? BTN_BG : NEUTRAL_BG, color: towerView === 'group' ? BTN_TEXT : NEUTRAL_TEXT }}>모임 전체 ({groupTower.length})</button>
        </div>
        {(() => {
          const list = towerView === 'mine' ? myTower : groupTower;
          if (list.length === 0) return <p className="text-xs text-center py-6" style={{ color: MUTE }}>아직 쌓인 책이 없어요.</p>;
          return (
            <div className="flex flex-col-reverse gap-1 max-h-96 overflow-y-auto pr-1">
              {list.map((t) => {
                const owner = towerView === 'group' ? members.find((m) => m.id === t.member_id) : null;
                return (
                  <div key={t.id} className="rounded-lg px-2.5 py-2 flex items-center justify-between" style={{ background: t.color, opacity: 0.92 }}>
                    <div className="min-w-0">
                      <div className="text-xs font-semibold truncate" style={{ color: '#1E1C16' }}>{t.book_title}{owner && ` · ${dispName(owner.name, isLoggedIn)}`}</div>
                      <div className="text-[10px]" style={{ color: 'rgba(30,28,22,0.7)', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(t.finished_date)}</div>
                    </div>
                    {towerView === 'mine' && currentMember && (
                      <button onClick={() => removeTowerEntry(t.id)} className="p-1" aria-label="책탑에서 제거"><X size={12} style={{ color: 'rgba(30,28,22,0.6)' }} /></button>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })()}
      </Card>

      {viewing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.85)' }} onClick={() => setViewingId(null)}>
          <div className="relative max-w-full max-h-full flex flex-col items-center" onClick={(e) => e.stopPropagation()}>
            {viewingIdx > 0 && (
              <button onClick={showPrev} className="absolute left-0 top-1/2 -translate-y-1/2 -translate-x-1 p-2 rounded-full" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }} aria-label="이전 사진"><ChevronLeft size={20} /></button>
            )}
            {viewingIdx < sorted.length - 1 && (
              <button onClick={showNext} className="absolute right-0 top-1/2 -translate-y-1/2 translate-x-1 p-2 rounded-full" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }} aria-label="다음 사진"><ChevronRight size={20} /></button>
            )}
            <img src={publicUrl('photos', viewing.file_path)} className="max-w-full max-h-[65vh] rounded-xl" alt="" />
            <div className="flex items-center gap-3 mt-3">
              <span className="text-sm" style={{ color: '#FFFFFF' }}>{dispName(viewing.uploader_name, isLoggedIn)} · {fmtDate(viewing.created_at)} {fmtTime(viewing.created_at)}</span>
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => { setEditingCaption(!editingCaption); setCaptionInput(viewing.caption || ''); }} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><FileText size={15} /></button>
              )}
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => { setEditingDate(!editingDate); setDateInput(viewing.created_at.slice(0, 10)); }} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }} aria-label="사진 날짜 수정"><Pencil size={15} /></button>
              )}
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => requestDelete(() => removePhoto(viewing), '이 사진을 삭제할까요?')} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#F0A87C' }} aria-label="사진 삭제"><Trash2 size={15} /></button>
              )}
              <button onClick={() => setViewingId(null)} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }} aria-label="닫기"><X size={15} /></button>
            </div>
            {viewing.caption && !editingCaption && (
              <p className="text-sm mt-2 max-w-sm text-center px-4" style={{ color: 'rgba(255,255,255,0.85)' }}>{viewing.caption}</p>
            )}
            {editingCaption && (
              <div className="flex items-center gap-2 mt-2 w-full max-w-sm px-4" onClick={(e) => e.stopPropagation()}>
                <input value={captionInput} onChange={(e) => setCaptionInput(e.target.value)} placeholder="캡션 추가" className="flex-1 rounded-lg border px-2 py-1.5 text-sm outline-none" style={{ background: '#17150F', borderColor: '#332F24', color: '#F2EEE3' }} aria-label="사진 캡션" />
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
    await reload(['meeting_locations']);
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
    await reload(['absence_excuses']);
  };
  const clearMyExcuse = async () => {
    if (!myExcuseToday) return;
    await deleteRow('absence_excuses', 'id', myExcuseToday.id);
    await reload(['absence_excuses']);
  };

  const startSession = async () => {
    await insertRow('sessions', { id: uid('s'), date: today, created_at: new Date().toISOString() });
    if (!calendarDays.some((d) => d.date === today && d.type === '독서일')) {
      await insertRow('calendar_days', { id: uid('cd'), date: today, type: '독서일' });
    }
    await reload(['sessions', 'calendar_days']);
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
    await reload(['checkins']);
  };
  const checkOut = async () => {
    if (!myCheckin) return;
    setLocChecking(true);
    const locResult = await getLocationStatus();
    setLocChecking(false);
    await updateRow('checkins', 'id', myCheckin.id, { check_out_at: new Date().toISOString(), checkout_loc_ok: locResult.ok });
    await reload(['checkins']);
  };
  const resetMyCheckin = async () => {
    if (!myCheckin) return;
    if (!window.confirm('오늘 체크인/체크아웃 기록을 초기화할까요?')) return;
    await deleteRow('checkins', 'id', myCheckin.id);
    await reload(['checkins']);
  };
  const todaysCheckins = session ? checkins.filter((c) => c.session_id === session.id) : [];
  const getCheckin = (memberId) => todaysCheckins.find((c) => c.member_id === memberId);
  const checkInMember = async (memberId) => {
    if (!session || getCheckin(memberId)) return;
    await insertRow('checkins', { id: uid('c'), session_id: session.id, member_id: memberId, check_in_at: new Date().toISOString(), check_out_at: null });
    await reload(['checkins']);
  };
  const checkOutMember = async (memberId) => {
    const c = getCheckin(memberId); if (!c || c.check_out_at) return;
    await updateRow('checkins', 'id', c.id, { check_out_at: new Date().toISOString() });
    await reload(['checkins']);
  };

  const toggleSelect = (id) => setSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const selectAll = () => setSelectedIds(members.map((m) => m.id));
  const clearSelect = () => setSelectedIds([]);

  const bulkCheckInSelected = async () => {
    if (!session || selectedIds.length === 0) return;
    const now = new Date().toISOString();
    const additions = selectedIds.filter((id) => !getCheckin(id)).map((id) => ({ id: uid('c'), session_id: session.id, member_id: id, check_in_at: now, check_out_at: null }));
    if (additions.length) await supabase.from('checkins').insert(additions);
    await reload(['checkins']);
  };
  const bulkCheckOutSelected = async () => {
    if (!session || selectedIds.length === 0) return;
    const now = new Date().toISOString();
    const ids = selectedIds.map((id) => getCheckin(id)).filter((c) => c && !c.check_out_at).map((c) => c.id);
    if (ids.length) await supabase.from('checkins').update({ check_out_at: now }).in('id', ids);
    await reload(['checkins']);
  };
  const bulkSetCheckInTime = async () => {
    if (!session || selectedIds.length === 0 || !timeInInput) return;
    const iso = new Date(`${today}T${timeInInput}:00`).toISOString();
    for (const id of selectedIds) {
      const c = getCheckin(id);
      if (c) await updateRow('checkins', 'id', c.id, { check_in_at: iso });
      else await insertRow('checkins', { id: uid('c'), session_id: session.id, member_id: id, check_in_at: iso, check_out_at: null });
    }
    await reload(['checkins']);
  };
  const bulkSetCheckOutTime = async () => {
    if (!session || selectedIds.length === 0 || !timeOutInput) return;
    const iso = new Date(`${today}T${timeOutInput}:00`).toISOString();
    for (const id of selectedIds) {
      const c = getCheckin(id);
      if (c) await updateRow('checkins', 'id', c.id, { check_out_at: iso });
    }
    await reload(['checkins']);
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
                    <button onClick={() => setMyExcuse('업무')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>업무</button>
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
                    {c && !c.check_out_at && <button onClick={() => checkOutMember(m.id)} className="p-1.5 rounded-lg" style={{ background: NEUTRAL_BG, color: '#F0A87C' }} aria-label={`${m.name} 체크아웃`}><LogOut size={14} /></button>}
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
              <input type="time" value={timeInInput} onChange={(e) => setTimeInInput(e.target.value)} className="flex-1 rounded-lg border px-2 py-1.5 text-xs outline-none" style={inputStyle} aria-label="일괄 체크인 시각" />
              <GhostBtn onClick={bulkSetCheckInTime} icon={LogIn}>체크인 시간 지정</GhostBtn>
            </div>
            <div className="flex items-center gap-2">
              <input type="time" value={timeOutInput} onChange={(e) => setTimeOutInput(e.target.value)} className="flex-1 rounded-lg border px-2 py-1.5 text-xs outline-none" style={inputStyle} aria-label="일괄 체크아웃 시각" />
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
                  <div className="flex items-center gap-2"><Stamp role={m?.role || '회원'} size={24} tilt={0} /><span style={{ color: INK }}>{m ? dispName(m.name, !!cur
