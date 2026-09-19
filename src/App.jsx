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
const BALLOON_REVEAL_PHOTO = 'data:image/jpeg;base64,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';
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
const TABLES = ['members', 'notices', 'notice_views', 'sessions', 'checkins', 'penalty_completions', 'calendar_days', 'settings', 'photos', 'absence_excuses', 'meeting_locations', 'dues_payments', 'expenses', 'dinner_collections', 'book_shares', 'notifications', 'book_tower_entries', 'birthday_balloons'];

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
  const [birthdayBalloons, setBirthdayBalloons] = useState([]);
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
      if (data.birthday_balloons) setBirthdayBalloons(data.birthday_balloons);
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

  // 오늘 생일인 멤버가 있으면 00시~24시 하루 종일 화면 전체에 꽃가루 효과
  const todayMdForConfetti = todayStr().slice(5, 10);
  const hasBirthdayToday = members.some((m) => m.birthday && mdOf(m.birthday) === todayMdForConfetti);
  const confettiColors = ['#F0A87C', '#7FA8D9', '#EFC94C', '#7FDCCF', '#E0958C', '#D9C24C'];
  const confettiPieces = useMemo(() => Array.from({ length: 50 }).map((_, i) => {
    const duration = 4.5 + Math.random() * 3.5; // 기존보다 약 1.2배 빠르게
    return {
      left: Math.round(Math.random() * 100),
      delay: -(Math.random() * duration).toFixed(2), // 음수 딜레이 - 처음부터 낙하 중간 지점에서 시작해 최상단에 쌓여 보이는 현상 방지
      duration: duration.toFixed(2),
      color: confettiColors[i % confettiColors.length],
      size: 6 + Math.round(Math.random() * 6),
      rotate: Math.round(Math.random() * 360),
    };
  }), [hasBirthdayToday]);

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
        @keyframes confettiFall {
          0% { transform: translateY(-5vh) rotate(0deg); opacity: 0.9; }
          100% { transform: translateY(105vh) rotate(720deg); opacity: 0.9; }
        }
      `}</style>
      {hasBirthdayToday && (
        <div className="fixed inset-0 pointer-events-none overflow-hidden" style={{ zIndex: 60 }} aria-hidden="true">
          {confettiPieces.map((p, i) => (
            <span key={i} style={{
              position: 'absolute', top: 0, left: `${p.left}%`, width: p.size, height: p.size * 0.4,
              background: p.color, borderRadius: 2, opacity: 0.9,
              animation: `confettiFall ${p.duration}s linear ${p.delay}s infinite`,
              transform: `rotate(${p.rotate}deg)`,
            }} />
          ))}
        </div>
      )}
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

        {tab === 'notice' && <NoticeScreen notices={notices} noticeViews={noticeViews} currentMember={currentMember} canManage={canManageUsers} reload={reload} members={members} requestDelete={requestDelete} birthdayBalloons={birthdayBalloons} />}
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

function NoticeScreen({ notices, noticeViews, currentMember, canManage, reload, members, requestDelete, birthdayBalloons }) {
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
  const todayFull = todayStr();
  const [sendingBalloon, setSendingBalloon] = useState(false);
  const BALLOON_COLORS = ['#F23B3B', '#F2871A', '#F5D400', '#3DBF54', '#2E86E0', '#8A4FE0', '#F0479C'];
  const todaysBalloons = (birthdayBalloons || []).filter((b) => b.for_date === todayFull).sort((a, b) => a.created_at.localeCompare(b.created_at));
  const sendBalloon = async () => {
    if (!currentMember || sendingBalloon) return;
    setSendingBalloon(true);
    const color = BALLOON_COLORS[Math.floor(Math.random() * BALLOON_COLORS.length)];
    await insertRow('birthday_balloons', { id: uid('bl'), for_date: todayFull, author_id: currentMember.id, author_name: currentMember.name, color, created_at: new Date().toISOString() });
    await reload();
    setSendingBalloon(false);
  };
  const removeBalloon = async (id) => { await deleteRow('birthday_balloons', 'id', id); await reload(); };
  const givenNameOnly = (name) => (name && name.length > 1 ? name.slice(1) : name); // 성 빼고 이름만
  const [flyingBalloonIds, setFlyingBalloonIds] = useState({});
  const [balloonsSettled, setBalloonsSettled] = useState(false);
  const [heartFormed, setHeartFormed] = useState(false); // 처음엔 흩어져 있다가, 한번 풍선을 띄우면 사진 주변으로 모임
  const [showBalloonPhoto, setShowBalloonPhoto] = useState(false); // 풍선이 다 날아가고 나면 사진이 등장
  const FLY_UP_DURATION = 3.0;
  const RETURN_DURATION = 2.8;
  const popBalloon = (id) => {
    setFlyingBalloonIds((prev) => ({ ...prev, [id]: true }));
    setHeartFormed(true); // 뜨는 순간부터 모두 사진 주변 자리를 향해 이동을 시작 - 돌아올 때 이미 그 영역 안에 도착해있도록
    setTimeout(() => setShowBalloonPhoto(true), FLY_UP_DURATION * 1000); // 풍선이 화면 위로 다 사라진 시점에 사진 등장
    setTimeout(() => {
      setFlyingBalloonIds((prev) => ({ ...prev, [id]: false }));
    }, (FLY_UP_DURATION + RETURN_DURATION) * 1000);
  };
  const shadeColor = (hex, percent) => {
    const num = parseInt(hex.replace('#', ''), 16);
    const clamp = (v) => Math.max(0, Math.min(255, v));
    const r = clamp((num >> 16) + Math.round(2.55 * percent));
    const g = clamp(((num >> 8) & 0xff) + Math.round(2.55 * percent));
    const bl = clamp((num & 0xff) + Math.round(2.55 * percent));
    return `#${((1 << 24) | (r << 16) | (g << 8) | bl).toString(16).slice(1)}`;
  };

  // 풍선이 쌓일 자리를 하트 모양으로 미리 계산해둠 — 바깥쪽 윤곽선부터, 그리고 각 레이어 안에서도 좌우 대칭 쌍으로 채워져서 몇 개만 있어도 하트 윤곽이 바로 보임
  const heartSlots = useMemo(() => {
    const layers = [
      { k: 1.00, n: 11 },
      { k: 0.86, n: 9 },
      { k: 0.70, n: 7 },
      { k: 0.52, n: 5 },
      { k: 0.32, n: 3 },
    ];
    // t=0(위 중앙 홈)을 시작으로, 좌우 대칭이 되는 인덱스 쌍을 번갈아 추가하는 순서
    const symmetricOrder = (n) => {
      const order = [0];
      for (let k = 1; k <= Math.floor(n / 2); k++) {
        order.push(k);
        if (n - k !== k && n - k !== 0) order.push(n - k);
      }
      return order;
    };
    const raw = [];
    layers.forEach(({ k, n }) => {
      symmetricOrder(n).forEach((i) => {
        const t = (i / n) * Math.PI * 2;
        const x = k * 16 * Math.pow(Math.sin(t), 3);
        const y = k * (13 * Math.cos(t) - 5 * Math.cos(2 * t) - 2 * Math.cos(3 * t) - Math.cos(4 * t));
        raw.push({ x, y });
      });
    });
    const xs = raw.map((p) => p.x); const ys = raw.map((p) => p.y);
    const minX = Math.min(...xs), maxX = Math.max(...xs), minY = Math.min(...ys), maxY = Math.max(...ys);
    return raw.map((p) => ({ left: ((p.x - minX) / (maxX - minX)) * 100, top: ((p.y - minY) / (maxY - minY)) * 100 }));
  }, []);

  // 풍선이 사진 주변을 둘러싸는 자리 (동심원 형태, 사진이 들어갈 가운데는 비워둠)
  const ringSlots = useMemo(() => {
    const rings = [
      { r: 36, n: 8 },
      { r: 45, n: 12 },
      { r: 53, n: 16 },
    ];
    const positions = [];
    rings.forEach(({ r, n }, ringIdx) => {
      for (let i = 0; i < n; i++) {
        const angle = (i / n) * Math.PI * 2 + ringIdx * 0.35;
        positions.push({ left: 50 + r * Math.cos(angle), top: 50 + r * 0.82 * Math.sin(angle) });
      }
    });
    return positions;
  }, []);

  // 접속 시 순차 등장 애니메이션이 끝나면, 이후로는 일반 흔들림만 반복하도록 전환
  useEffect(() => {
    setBalloonsSettled(false);
    const maxDelay = Math.min(Math.max(todaysBalloons.length - 1, 0), 10) * 0.09 + 0.7;
    const t = setTimeout(() => setBalloonsSettled(true), (maxDelay + 0.15) * 1000);
    return () => clearTimeout(t);
  }, [todaysBalloons.length]);

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
      <style>{`
        @keyframes balloonRiseIn {
          0% { transform: translateY(90px) scale(0.5); opacity: 0; }
          65% { transform: translateY(-8px) scale(1.06); opacity: 1; }
          100% { transform: translateY(0) scale(1); opacity: 1; }
        }
        @keyframes balloonBob {
          0% { transform: translateY(0) rotate(0deg); }
          12.5% { transform: translateY(-1px) rotate(-3deg); }
          25% { transform: translateY(-2.5px) rotate(-6deg); }
          37.5% { transform: translateY(-4px) rotate(-3deg); }
          50% { transform: translateY(-5px) rotate(0deg); }
          62.5% { transform: translateY(-4px) rotate(3deg); }
          75% { transform: translateY(-2.5px) rotate(6deg); }
          87.5% { transform: translateY(-1px) rotate(3deg); }
          100% { transform: translateY(0) rotate(0deg); }
        }
        @keyframes balloonFlyUp {
          0% { transform: translate(0, 0) rotate(0deg) scale(1); opacity: 1; }
          25% { transform: translate(10px, -35vh) rotate(-6deg) scale(0.92); opacity: 1; }
          50% { transform: translate(-12px, -70vh) rotate(5deg) scale(0.8); opacity: 1; }
          75% { transform: translate(10px, -105vh) rotate(-5deg) scale(0.65); opacity: 0.9; }
          100% { transform: translate(-8px, -140vh) rotate(4deg) scale(0.5); opacity: 0; }
        }
        @keyframes balloonReturnFromBottom {
          0% { transform: translate(-8px, 140vh) scale(0.5); opacity: 0; }
          25% { transform: translate(10px, 105vh) scale(0.65); opacity: 0.9; }
          50% { transform: translate(-12px, 70vh) scale(0.8); opacity: 1; }
          75% { transform: translate(10px, 35vh) scale(0.92); opacity: 1; }
          100% { transform: translate(0, 0) scale(1); opacity: 1; }
        }
        @keyframes photoRevealIn {
          0% { opacity: 0; transform: translate(-50%, -50%) scale(0.4); }
          70% { opacity: 1; transform: translate(-50%, -50%) scale(1.08); }
          100% { opacity: 1; transform: translate(-50%, -50%) scale(1); }
        }
      `}</style>
      {birthdayFolksToday.length > 0 && (
        <Card className="text-center">
          <div className="flex items-center justify-center gap-2 mb-1">
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
            <span className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>오늘은 {birthdayFolksToday.map((m) => dispName(m.name, isLoggedIn)).join(', ')}님 생일이에요!</span>
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
          </div>
          <p className="text-sm" style={{ color: MUTE }}>생일 축하해요~ 행복한 하루 되세요 🎂</p>
          {todaysBalloons.length > 0 && (
            <div className="relative mx-auto mt-3" style={{ width: '100%', maxWidth: 360, height: 320 }}>
              {todaysBalloons.map((b, i) => {
                const wrapIdx = i % ringSlots.length;
                const wrapRound = Math.floor(i / ringSlots.length);
                const slot = ringSlots[wrapIdx];
                // 슬롯을 다 채우고 넘치면, id 기반 미세한 오프셋을 줘서 겹치지 않게
                const jHash = b.id.split('').reduce((s, c) => s + c.charCodeAt(0), 0);
                const jx = wrapRound ? ((jHash % 7) - 3) * 1.2 : 0;
                const jy = wrapRound ? (((jHash >> 3) % 7) - 3) * 1.2 : 0;
                // 하트 형성 전에는 영역 안에 각자 흩어진 자리에 있음 (id 기반이라 매번 렌더링해도 위치 고정)
                const scatterLeft = 12 + (jHash % 76);
                const scatterTop = 12 + ((jHash >> 4) % 76);
                const posLeft = heartFormed ? slot.left + jx : scatterLeft;
                const posTop = heartFormed ? slot.top + jy : scatterTop;
                const gradId = `balloon-grad-${b.id}`;
                const riseDelay = Math.min(i, 10) * 0.09; // 접속 시 하나씩 순차적으로 올라오는 느낌
                const dur = 2.6 + (i % 3) * 0.4;
                const isFlying = !!flyingBalloonIds[b.id];
                const anim = isFlying
                  ? `balloonFlyUp ${FLY_UP_DURATION}s linear forwards, balloonReturnFromBottom ${RETURN_DURATION}s linear ${FLY_UP_DURATION}s both`
                  : balloonsSettled
                    ? `balloonBob ${dur}s linear infinite`
                    : `balloonRiseIn 0.7s ease-out ${riseDelay}s both, balloonBob ${dur}s linear ${riseDelay + 0.7}s infinite`;
                return (
                  <div key={b.id} className="absolute" style={{ left: `${posLeft}%`, top: `${posTop}%`, transform: 'translate(-50%, -50%)', transition: 'left 1.3s ease-in-out, top 1.3s ease-in-out' }}>
                    <div onClick={() => popBalloon(b.id)} className="relative flex flex-col items-center cursor-pointer" style={{ width: 41, transformOrigin: '50% 100%', animation: anim }}>
                      <div className="relative" style={{ width: 39, height: 36 }}>
                        <svg width="39" height="36" viewBox="2.25 3 19.5 18" style={{ filter: 'drop-shadow(0 3px 3px rgba(0,0,0,0.35)) saturate(1.15)', overflow: 'visible', display: 'block' }}>
                          <defs>
                            <radialGradient id={gradId} cx="32%" cy="26%" r="80%">
                              <stop offset="0%" stopColor={shadeColor(b.color, 55)} />
                              <stop offset="55%" stopColor={b.color} />
                              <stop offset="100%" stopColor={shadeColor(b.color, -28)} />
                            </radialGradient>
                          </defs>
                          <path d="M11.645 20.91l-.007-.003-.022-.012a15.247 15.247 0 01-.383-.218 25.18 25.18 0 01-4.244-3.17C4.688 15.36 2.25 12.174 2.25 8.25 2.25 5.322 4.714 3 7.688 3A5.5 5.5 0 0112 5.052 5.5 5.5 0 0116.313 3c2.973 0 5.437 2.322 5.437 5.25 0 3.925-2.438 7.111-4.739 9.256a25.175 25.175 0 01-4.244 3.17 15.247 15.247 0 01-.383.219l-.022.012-.007.004-.003.001a.752.752 0 01-.704 0l-.003-.001z" fill={`url(#${gradId})`} stroke={shadeColor(b.color, -35)} strokeWidth="0.3" strokeLinejoin="round" />
                          {/* 하이라이트 - 유광 반사 느낌 (회전 없이 단순한 원으로) */}
                          <circle cx="8" cy="8.3" r="1.7" fill="rgba(255,255,255,0.7)" />
                        </svg>
                        {/* 이름표 - 성 빼고 이름만, 풍선 한가운데 */}
                        <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
                          <span className="text-[8px] font-bold leading-none whitespace-nowrap" style={{ color: '#FFFFFF', textShadow: '0 1px 2px rgba(0,0,0,0.45)' }}>{givenNameOnly(dispName(b.author_name, isLoggedIn))}</span>
                        </div>
                      </div>
                      {/* 끈 - 하트 바로 아래, 풍선 세로 길이의 절반 정도 */}
                      <svg width="8" height="16" viewBox="0 0 8 16" style={{ overflow: 'visible', display: 'block', marginTop: -1 }}>
                        <path d="M4 0 C 6 4.5, 2 11.5, 4 16" stroke="rgba(255,255,255,0.45)" strokeWidth="1" fill="none" />
                      </svg>
                      {currentMember?.id === b.author_id && (
                        <button onClick={(e) => { e.stopPropagation(); requestDelete(() => removeBalloon(b.id), '이 풍선을 없앨까요?'); }} className="absolute -top-1.5 -right-1.5 rounded-full p-0.5" style={{ background: CARD_BG }} aria-label="풍선 삭제"><X size={10} style={{ color: MUTE }} /></button>
                      )}
                    </div>
                  </div>
                );
              })}
              {showBalloonPhoto && (
                <div className="absolute" style={{ left: '50%', top: '50%', animation: 'photoRevealIn 0.7s ease-out both', pointerEvents: 'none' }}>
                  <img src={BALLOON_REVEAL_PHOTO} alt="" style={{ width: 216, height: 216, borderRadius: 28, objectFit: 'cover', border: '3px solid rgba(255,255,255,0.85)', boxShadow: '0 4px 14px rgba(0,0,0,0.45)' }} />
                </div>
              )}
            </div>
          )}
          {currentMember && (
            <button onClick={sendBalloon} disabled={sendingBalloon} className="mt-3 rounded-full px-4 py-2 text-xs font-semibold disabled:opacity-50" style={{ background: BTN_BG, color: BTN_TEXT }}>🎈 축하 풍선 띄우기</button>
          )}
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
  const [borrowedDateInput, setBorrowedDateInput] = useState('');
  const [editingShare, setEditingShare] = useState(false);
  const [editShareTitle, setEditShareTitle] = useState('');
  const [editShareAuthor, setEditShareAuthor] = useState('');
  const [editSharePublisher, setEditSharePublisher] = useState('');
  const [coverCache, setCoverCache] = useState({}); // { [shareId]: 'loading' | url | 'none' }
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
  const saveShareEdit = async (share) => {
    if (!editShareTitle.trim()) return;
    await updateRow('book_shares', 'id', share.id, { book_title: editShareTitle.trim(), book_author: editShareAuthor.trim() || null, book_publisher: editSharePublisher.trim() || null });
    setEditingShare(false);
    await reload();
  };
  // 1단계: 요청하기 — 아직 확정 아님, 글쓴이가 확인 후 확정해야 함
  const requestShare = async (share) => {
    if (!currentMember) return;
    await updateRow('book_shares', 'id', share.id, { status: 'requested', matched_by: currentMember.id });
    await notify(share.posted_by, `${currentMember.name}님이 [${share.book_title}] ${share.kind === 'offer' ? '제공' : '요청'}에 응답했어요. 확인 후 확정해주세요.`, share.id);
    await reload();
  };
  // 요청 취소 — 글쓴이가 다시 대기 상태로 되돌림 (다른 사람이 다시 요청할 수 있도록)
  const cancelRequest = async (share) => {
    await updateRow('book_shares', 'id', share.id, { status: 'open', matched_by: null });
    await reload();
  };
  // 2단계: 확정 — 이때 대여일/반납기한이 정해짐
  const confirmMatch = async (share) => {
    const today = todayStr();
    const due = new Date(); due.setDate(due.getDate() + 14);
    const dueStr = `${due.getFullYear()}-${pad(due.getMonth() + 1)}-${pad(due.getDate())}`;
    await updateRow('book_shares', 'id', share.id, { status: 'matched', borrowed_at: today, due_date: dueStr });
    await notify(share.matched_by, `[${share.book_title}] 대여가 확정됐어요. 반납기한: ${dueStr}`, share.id);
    await reload();
  };
  const updateDueDate = async (share, newDate) => {
    await updateRow('book_shares', 'id', share.id, { due_date: newDate });
    await reload();
  };
  const updateBorrowedDate = async (share, newDate) => {
    const due = new Date(`${newDate}T00:00:00`); due.setDate(due.getDate() + 14);
    const dueStr = `${due.getFullYear()}-${pad(due.getMonth() + 1)}-${pad(due.getDate())}`;
    await updateRow('book_shares', 'id', share.id, { borrowed_at: newDate, due_date: dueStr });
    setDueDateInput('');
    await reload();
  };
  const randomPastel = () => {
    // 레퍼런스처럼 채도 낮은 차분한 톤(크림, 모브, 브라운, 더스티핑크 등)
    const palette = ['#D9CFC1', '#B9A695', '#D98C8C', '#3A2C25', '#D6C79E', '#C9B8AC'];
    return palette[Math.floor(Math.random() * palette.length)];
  };
  const isDarkColor = (hex) => {
    const c = (hex || '#D9CFC1').replace('#', '');
    const r = parseInt(c.substr(0, 2), 16), g = parseInt(c.substr(2, 2), 16), b = parseInt(c.substr(4, 2), 16);
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5;
  };
  const confirmReturn = async (share) => {
    const borrowerId = share.kind === 'offer' ? share.matched_by : share.posted_by;
    const today = todayStr();
    await updateRow('book_shares', 'id', share.id, { status: 'returned', returned_at: today });
    await insertRow('book_tower_entries', { id: uid('bt'), member_id: borrowerId, book_title: share.book_title, start_date: share.borrowed_at, finished_date: today, current_page: null, color: randomPastel(), source_share_id: share.id, created_at: new Date().toISOString() });
    await notify(borrowerId, `[${share.book_title}] 반납 완료 처리됐어요. 내 책탑에 추가됐어요.`, share.id);
    await reload();
  };
  const deleteBookShare = async (share) => {
    await deleteRow('book_shares', 'id', share.id);
    setViewingShareId(null);
    await reload();
  };
  const shareStatusInfo = (s) => {
    if (s.status === 'open') return { label: s.kind === 'offer' ? '대여가능' : '대기중', style: { background: '#12302C', color: '#7FDCCF' } };
    if (s.status === 'requested') return { label: '요청중', style: { background: '#332815', color: '#EFC94C' } };
    if (s.status === 'matched') return { label: '대여중', style: { background: '#3A2E10', color: '#EFC94C' } };
    return { label: '반납완료', style: { background: NEUTRAL_BG, color: MUTE, border: `1px solid ${LINE}` } };
  };
  const offers = bookShares.filter((s) => s.kind === 'offer').sort((a, b) => b.created_at.localeCompare(a.created_at));
  const requests = bookShares.filter((s) => s.kind === 'request').sort((a, b) => b.created_at.localeCompare(a.created_at));
  const viewingShare = bookShares.find((s) => s.id === viewingShareId) || null;

  useEffect(() => {
    if (!viewingShare) return;
    if (coverCache[viewingShare.id]) return; // 이미 조회했으면 재요청 안 함
    let cancelled = false;
    const key = viewingShare.id;
    setCoverCache((prev) => ({ ...prev, [key]: 'loading' }));
    const q = [viewingShare.book_title, viewingShare.book_author].filter(Boolean).map((v) => v.trim()).filter(Boolean);
    const query = encodeURIComponent(`intitle:${q[0] || ''}${q[1] ? ` inauthor:${q[1]}` : ''}`);
    fetch(`https://www.googleapis.com/books/v1/volumes?q=${query}&maxResults=1&country=KR`)
      .then((res) => (res.ok ? res.json() : null))
      .then((data) => {
        if (cancelled) return;
        const thumb = data?.items?.[0]?.volumeInfo?.imageLinks?.thumbnail || data?.items?.[0]?.volumeInfo?.imageLinks?.smallThumbnail || null;
        const url = thumb ? thumb.replace(/^http:/, 'https:') : 'none';
        setCoverCache((prev) => ({ ...prev, [key]: url }));
      })
      .catch(() => { if (!cancelled) setCoverCache((prev) => ({ ...prev, [key]: 'none' })); });
    return () => { cancelled = true; };
  }, [viewingShare?.id]);

  // ---------- 책탑 ----------
  const [towerView, setTowerView] = useState('mine'); // 'mine' | 'group'
  const [showTowerAdd, setShowTowerAdd] = useState(false);
  const [towerTitleInput, setTowerTitleInput] = useState('');
  const [towerStartInput, setTowerStartInput] = useState(todayStr());
  const [towerPageInput, setTowerPageInput] = useState('');
  const [towerFinishedInput, setTowerFinishedInput] = useState('');
  const [editingTowerId, setEditingTowerId] = useState(null);
  const addManualTowerEntry = async () => {
    if (!currentMember || !towerTitleInput.trim()) return;
    await insertRow('book_tower_entries', {
      id: uid('bt'), member_id: currentMember.id, book_title: towerTitleInput.trim(),
      start_date: towerStartInput || null, current_page: towerPageInput ? parseInt(towerPageInput, 10) : null, finished_date: towerFinishedInput || null,
      color: randomPastel(), source_share_id: null, created_at: new Date().toISOString(),
    });
    setTowerTitleInput(''); setTowerStartInput(todayStr()); setTowerPageInput(''); setTowerFinishedInput(''); setShowTowerAdd(false);
    await reload();
  };
  const saveTowerEdit = async (entry) => {
    await updateRow('book_tower_entries', 'id', entry.id, { start_date: towerStartInput || null, current_page: towerPageInput ? parseInt(towerPageInput, 10) : null, finished_date: towerFinishedInput || null });
    setEditingTowerId(null);
    await reload();
  };
  const startTowerEdit = (entry) => {
    setEditingTowerId(entry.id);
    setTowerStartInput(entry.start_date || '');
    setTowerPageInput(entry.current_page ? String(entry.current_page) : '');
    setTowerFinishedInput(entry.finished_date || '');
  };
  const removeTowerEntry = async (entryId) => { await deleteRow('book_tower_entries', 'id', entryId); await reload(); };
  const towerSortKey = (t) => t.finished_date || t.start_date || t.created_at.slice(0, 10);
  const myTower = currentMember ? bookTowerEntries.filter((t) => t.member_id === currentMember.id).sort((a, b) => towerSortKey(a).localeCompare(towerSortKey(b))) : [];
  const groupTower = [...bookTowerEntries].sort((a, b) => towerSortKey(a).localeCompare(towerSortKey(b)));

  return (
    <div className="space-y-4">
      <Card>
        <div className="flex items-center gap-1.5 text-sm font-semibold mb-3" style={{ color: INK }}><ImageIcon size={16} style={{ color: '#7FDCCF' }} /> 포토로그</div>
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
            className="w-full rounded-xl border px-3 py-2.5 text-sm outline-none mt-3" style={inputStyle} />
        )}
        <div className="mt-3">
          {sorted.length === 0 ? (
            <p className="text-sm text-center py-6" style={{ color: MUTE }}>아직 올라온 사진이 없어요.</p>
          ) : filtered.length === 0 ? (
            <p className="text-sm text-center py-6" style={{ color: MUTE }}>검색 결과가 없어요.</p>
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
        </div>
        {filtered.length > PHOTOS_PER_PAGE && (
          <div className="flex items-center justify-center gap-3 mt-3">
            <button onClick={() => setPhotoPage((p) => Math.max(0, p - 1))} disabled={safePage === 0} aria-label="이전 페이지" className="p-1.5 rounded-full" style={{ color: safePage === 0 ? LINE : MUTE }}><ChevronLeft size={16} /></button>
            <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{safePage + 1} / {totalPhotoPages}</span>
            <button onClick={() => setPhotoPage((p) => Math.min(totalPhotoPages - 1, p + 1))} disabled={safePage >= totalPhotoPages - 1} aria-label="다음 페이지" className="p-1.5 rounded-full" style={{ color: safePage >= totalPhotoPages - 1 ? LINE : MUTE }}><ChevronRight size={16} /></button>
          </div>
        )}
      </Card>

      {/* ---------- 도서 공유함 ---------- */}
      <Card>
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-1.5 text-sm font-semibold" style={{ color: INK }}><BookOpen size={16} style={{ color: '#7FA8D9' }} /> 도서 공유함</div>
          {currentMember && <button onClick={() => setShowShareForm((v) => !v)} className="text-xs rounded-full px-3 py-1.5 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>{showShareForm ? '취소' : '글쓰기'}</button>}
        </div>
        {showShareForm && (
          <div className="space-y-2 mb-3 pb-3" style={{ borderBottom: `1px solid ${ROW_LINE}` }}>
            <div className="flex gap-2">
              <button onClick={() => setShareKind('offer')} className="flex-1 rounded-xl py-2 text-xs font-semibold" style={{ background: shareKind === 'offer' ? BTN_BG : NEUTRAL_BG, color: shareKind === 'offer' ? BTN_TEXT : NEUTRAL_TEXT }}>빌려줄까요?</button>
              <button onClick={() => setShareKind('request')} className="flex-1 rounded-xl py-2 text-xs font-semibold" style={{ background: shareKind === 'request' ? BTN_BG : NEUTRAL_BG, color: shareKind === 'request' ? BTN_TEXT : NEUTRAL_TEXT }}>빌려주실수있나요?</button>
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
        <div className="mb-1.5 text-xs font-semibold" style={{ color: MUTE }}>📚 빌려줄까요? ({offers.length})</div>
        <div className="space-y-1.5 mb-3">
          {offers.length === 0 && <p className="text-xs" style={{ color: MUTE }}>등록된 글이 없어요.</p>}
          {offers.map((s) => {
            const poster = members.find((m) => m.id === s.posted_by);
            return (
              <button key={s.id} onClick={() => { setViewingShareId(s.id); setDueDateInput(''); setBorrowedDateInput(''); setEditingShare(false); }} className="w-full flex items-center justify-between rounded-xl px-3 py-2 text-left" style={{ background: NEUTRAL_BG }}>
                <div className="min-w-0">
                  <div className="text-sm truncate" style={{ color: INK }}>{s.book_title}</div>
                  <div className="text-[10px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{dispName(poster?.name || '', isLoggedIn)}</div>
                </div>
                <span className="text-[10px] rounded-full px-2 py-0.5 font-semibold shrink-0 ml-2" style={shareStatusInfo(s).style}>{shareStatusInfo(s).label}</span>
              </button>
            );
          })}
        </div>
        <div className="mb-1.5 text-xs font-semibold" style={{ color: MUTE }}>🙋 빌려주실수있나요? ({requests.length})</div>
        <div className="space-y-1.5">
          {requests.length === 0 && <p className="text-xs" style={{ color: MUTE }}>등록된 글이 없어요.</p>}
          {requests.map((s) => {
            const poster = members.find((m) => m.id === s.posted_by);
            return (
              <button key={s.id} onClick={() => { setViewingShareId(s.id); setDueDateInput(''); setBorrowedDateInput(''); setEditingShare(false); }} className="w-full flex items-center justify-between rounded-xl px-3 py-2 text-left" style={{ background: NEUTRAL_BG }}>
                <div className="min-w-0">
                  <div className="text-sm truncate" style={{ color: INK }}>{s.book_title}</div>
                  <div className="text-[10px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{dispName(poster?.name || '', isLoggedIn)}</div>
                </div>
                <span className="text-[10px] rounded-full px-2 py-0.5 font-semibold shrink-0 ml-2" style={shareStatusInfo(s).style}>{shareStatusInfo(s).label}</span>
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
        const canEditPost = currentMember?.id === viewingShare.posted_by || canManage;
        return (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.7)' }} onClick={() => setViewingShareId(null)}>
            <div className="w-full max-w-sm rounded-2xl border p-5" style={{ background: CARD_BG, borderColor: LINE }} onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between mb-3">
                <span className="text-[11px] rounded-full px-2 py-0.5 font-semibold" style={{ background: viewingShare.kind === 'offer' ? '#1E2A38' : '#332815', color: viewingShare.kind === 'offer' ? '#7FA8D9' : '#EFC94C' }}>{viewingShare.kind === 'offer' ? '제공' : '요청'}</span>
                <div className="flex items-center gap-2">
                  {canEditPost && !editingShare && (
                    <button onClick={() => { setEditingShare(true); setEditShareTitle(viewingShare.book_title); setEditShareAuthor(viewingShare.book_author || ''); setEditSharePublisher(viewingShare.book_publisher || ''); }} aria-label="글 수정"><Pencil size={14} style={{ color: MUTE }} /></button>
                  )}
                  <button onClick={() => setViewingShareId(null)} aria-label="닫기"><X size={16} style={{ color: MUTE }} /></button>
                </div>
              </div>
              {editingShare ? (
                <div className="space-y-2 mb-3">
                  <input value={editShareTitle} onChange={(e) => setEditShareTitle(e.target.value)} placeholder="책 제목" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                  <input value={editShareAuthor} onChange={(e) => setEditShareAuthor(e.target.value)} placeholder="저자 (선택)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                  <input value={editSharePublisher} onChange={(e) => setEditSharePublisher(e.target.value)} placeholder="출판사 (선택)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                  <div className="flex gap-2">
                    <PrimaryBtn onClick={() => saveShareEdit(viewingShare)} icon={Check}>저장</PrimaryBtn>
                    <GhostBtn onClick={() => setEditingShare(false)}>취소</GhostBtn>
                  </div>
                </div>
              ) : (
                <>
                  {coverCache[viewingShare.id] && coverCache[viewingShare.id] !== 'none' && coverCache[viewingShare.id] !== 'loading' && (
                    <div className="flex justify-center mb-3">
                      <img src={coverCache[viewingShare.id]} alt="" className="rounded-lg shadow-md" style={{ height: 128, width: 'auto' }} />
                    </div>
                  )}
                  {coverCache[viewingShare.id] === 'loading' && (
                    <div className="flex justify-center mb-3">
                      <div className="rounded-lg animate-pulse" style={{ height: 128, width: 88, background: NEUTRAL_BG }} />
                    </div>
                  )}
                  <div className="text-base font-semibold mb-2" style={{ color: INK }}>{viewingShare.book_title}</div>
                  <div className="text-xs space-y-1 mb-3" style={{ color: NEUTRAL_TEXT }}>
                    {viewingShare.book_author && <div>저자: {viewingShare.book_author}</div>}
                    {viewingShare.book_publisher && <div>출판사: {viewingShare.book_publisher}</div>}
                    <div style={{ color: MUTE }}>작성자: {dispName(poster?.name || '', isLoggedIn)}</div>
                  </div>
                </>
              )}
              {viewingShare.status === 'open' && !editingShare && (
                <>
                  {currentMember && currentMember.id !== viewingShare.posted_by ? (
                    <PrimaryBtn onClick={() => { requestShare(viewingShare); setViewingShareId(null); }} icon={Check}>{viewingShare.kind === 'offer' ? '제가 빌릴게요' : '제가 빌려드릴게요'}</PrimaryBtn>
                  ) : currentMember?.id === viewingShare.posted_by ? (
                    <p className="text-xs" style={{ color: MUTE }}>다른 회원의 응답을 기다리는 중이에요.</p>
                  ) : null}
                </>
              )}
              {viewingShare.status === 'requested' && !editingShare && (
                <div className="space-y-2 pt-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="text-xs" style={{ color: NEUTRAL_TEXT }}>응답: {dispName(matcher?.name || '', isLoggedIn)}</div>
                  {currentMember?.id === viewingShare.posted_by ? (
                    <div className="flex gap-2">
                      <PrimaryBtn onClick={() => { confirmMatch(viewingShare); setViewingShareId(null); }} icon={Check}>확정하기</PrimaryBtn>
                      <GhostBtn onClick={() => cancelRequest(viewingShare)}>취소</GhostBtn>
                    </div>
                  ) : (
                    <p className="text-xs" style={{ color: MUTE }}>작성자가 확인 후 확정하면 대여가 시작돼요.</p>
                  )}
                </div>
              )}
              {(viewingShare.status === 'matched' || viewingShare.status === 'returned') && !editingShare && (
                <div className="space-y-2 pt-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="text-xs" style={{ color: NEUTRAL_TEXT }}>응답: {dispName(matcher?.name || '', isLoggedIn)}</div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs" style={{ color: MUTE }}>대여일:</span>
                    {viewingShare.status === 'matched' && (isOwner || currentMember?.id === borrowerId || canManage) ? (
                      <input type="date" value={borrowedDateInput || viewingShare.borrowed_at || ''} onChange={(e) => setBorrowedDateInput(e.target.value)} onBlur={() => borrowedDateInput && updateBorrowedDate(viewingShare, borrowedDateInput)}
                        className="rounded-lg border px-1.5 py-1 text-[11px] outline-none" style={inputStyle} aria-label="대여일" />
                    ) : (
                      <span className="text-xs" style={{ color: NEUTRAL_TEXT }}>{fmtDate(viewingShare.borrowed_at)}</span>
                    )}
                  </div>
                  {viewingShare.status === 'matched' && (
                    <div className="flex items-center gap-1.5">
                      <span className="text-xs" style={{ color: MUTE }}>반납기한:</span>
                      {isOwner || currentMember?.id === borrowerId || canManage ? (
                        <input type="date" value={dueDateInput || viewingShare.due_date || ''} onChange={(e) => setDueDateInput(e.target.value)} onBlur={() => dueDateInput && updateDueDate(viewingShare, dueDateInput)}
                          className="rounded-lg border px-1.5 py-1 text-[11px] outline-none" style={inputStyle} aria-label="반납기한" />
                      ) : (
                        <span className="text-xs" style={{ color: NEUTRAL_TEXT }}>{fmtDate(viewingShare.due_date)}</span>
                      )}
                    </div>
                  )}
                  {viewingShare.status === 'returned' && <div className="text-xs" style={{ color: '#7FDCCF' }}>✓ 반납완료 · {fmtDate(viewingShare.returned_at)}</div>}
                  {viewingShare.status === 'matched' && (isOwner || canManage) && (
                    <PrimaryBtn onClick={() => { confirmReturn(viewingShare); setViewingShareId(null); }} icon={Check}>반납 완료 처리</PrimaryBtn>
                  )}
                </div>
              )}
              {canEditPost && !editingShare && (
                <button onClick={() => requestDelete(() => deleteBookShare(viewingShare), '이 게시글을 삭제할까요? 대여 기록도 함께 사라져요.')} className="text-[11px] underline underline-offset-2 mt-3" style={{ color: '#F0A87C' }}>글 삭제</button>
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
            <input value={towerTitleInput} onChange={(e) => setTowerTitleInput(e.target.value)} placeholder="책 제목" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <div className="grid grid-cols-2 gap-2">
              <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>읽기 시작일</div><input type="date" value={towerStartInput} onChange={(e) => setTowerStartInput(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
              <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>다 읽은 날 (선택)</div><input type="date" value={towerFinishedInput} onChange={(e) => setTowerFinishedInput(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            </div>
            <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>현재 읽고 있는 페이지 (선택)</div><input type="number" value={towerPageInput} onChange={(e) => setTowerPageInput(e.target.value)} placeholder="예: 128" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
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
            <div className="flex flex-col-reverse gap-1.5 max-h-[28rem] overflow-y-auto pr-1">
              {list.map((t) => {
                const owner = towerView === 'group' ? members.find((m) => m.id === t.member_id) : null;
                const isMine = towerView === 'mine' && currentMember;
                const isEditing = editingTowerId === t.id;
                const dark = isDarkColor(t.color);
                const fg = dark ? '#F2EEE3' : '#2A2620';
                const fgMute = dark ? 'rgba(242,238,227,0.68)' : 'rgba(42,38,32,0.62)';
                const statusText = t.finished_date ? '완독' : t.current_page ? `p.${t.current_page}` : '읽는 중';
                // id를 기반으로 한 안정적인 값으로 살짝 다른 두께를 줘서 실제 책처럼 자연스럽게
                const hash = t.id.split('').reduce((s, c) => s + c.charCodeAt(0), 0);
                const containerH = [64, 70, 76][hash % 3];
                const lineShade = dark ? 'rgba(255,255,255,0.16)' : 'rgba(0,0,0,0.12)';
                const highlight = dark ? 'rgba(255,255,255,0.06)' : 'rgba(255,255,255,0.35)';
                const shade = dark ? 'rgba(0,0,0,0.22)' : 'rgba(0,0,0,0.06)';
                return (
                  <div key={t.id} className="relative overflow-hidden" style={{
                    height: containerH, borderRadius: 14,
                    background: `linear-gradient(180deg, ${highlight} 0%, transparent 25%, transparent 75%, ${shade} 100%), ${t.color}`,
                    boxShadow: '0 2px 5px rgba(0,0,0,0.3)',
                  }}>
                    {/* 왼쪽 페이지 결 — 책이 옆으로 누워 여러 장 겹친 느낌 */}
                    <div className="absolute pointer-events-none" style={{ top: 6, bottom: 6, left: 12, width: 2, background: lineShade, borderRadius: 1 }} />
                    <div className="absolute pointer-events-none" style={{ top: 6, bottom: 6, left: 17, width: 2, background: lineShade, borderRadius: 1 }} />
                    {/* 우상단 책갈피 리본 */}
                    <div className="absolute pointer-events-none" style={{ top: 0, right: 20, width: 11, height: 22, background: dark ? 'rgba(255,255,255,0.28)' : 'rgba(0,0,0,0.2)', clipPath: 'polygon(0 0, 100% 0, 100% 100%, 50% 72%, 0 100%)' }} />
                    {isEditing ? (
                      <div className="absolute inset-0 flex flex-col justify-center gap-1.5 px-6" style={{ background: CARD_BG }}>
                        <div className="text-sm font-bold truncate" style={{ color: INK }}>{t.book_title}</div>
                        <div className="grid grid-cols-2 gap-1.5">
                          <input type="date" value={towerStartInput} onChange={(e) => setTowerStartInput(e.target.value)} className="rounded-lg px-2 py-1 text-[11px] outline-none" style={inputStyle} aria-label="읽기 시작일" />
                          <input type="date" value={towerFinishedInput} onChange={(e) => setTowerFinishedInput(e.target.value)} className="rounded-lg px-2 py-1 text-[11px] outline-none" style={inputStyle} aria-label="다 읽은 날" />
                        </div>
                        <div className="flex gap-1.5 items-center">
                          <input type="number" value={towerPageInput} onChange={(e) => setTowerPageInput(e.target.value)} placeholder="현재 페이지" className="flex-1 rounded-lg px-2 py-1 text-[11px] outline-none" style={inputStyle} />
                          <button onClick={() => saveTowerEdit(t)} className="text-[11px] rounded-full px-2.5 py-1 font-semibold shrink-0" style={{ background: '#1E1C16', color: '#F2EEE3' }}>저장</button>
                          <button onClick={() => setEditingTowerId(null)} className="text-[11px] rounded-full px-2.5 py-1 font-semibold shrink-0" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>취소</button>
                        </div>
                      </div>
                    ) : (
                      <div className="relative flex items-center justify-between h-full pl-7 pr-4 gap-2">
                        <div className="min-w-0">
                          <div className="text-sm font-bold truncate" style={{ color: fg }}>{t.book_title}</div>
                          <div className="text-[11px] truncate" style={{ color: fgMute }}>
                            {owner ? dispName(owner.name, isLoggedIn) : (t.start_date ? `시작 ${fmtDate(t.start_date)}` : '')}
                          </div>
                        </div>
                        <div className="flex items-center gap-1.5 shrink-0">
                          <span className="text-xs font-medium" style={{ color: fgMute }}>{statusText}</span>
                          {isMine && (
                            <>
                              <button onClick={() => startTowerEdit(t)} className="p-1" aria-label="책탑 항목 수정"><Pencil size={12} style={{ color: fgMute }} /></button>
                              <button onClick={() => removeTowerEntry(t.id)} className="p-1" aria-label="책탑에서 제거"><X size={12} style={{ color: fgMute }} /></button>
                            </>
                          )}
                        </div>
                      </div>
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
  const [expandedPenaltyRow, setExpandedPenaltyRow] = useState(null); // `${weekKey}_${memberId}` — 펼쳐진 벌칙 수정 패널
  const [penaltyDateDrafts, setPenaltyDateDrafts] = useState({}); // { `${weekKey}_${memberId}`: 'YYYY-MM-DD' }
  // 벌칙 수행 예정일만 저장 — confirmed는 건드리지 않음(날짜만 넣는다고 자동으로 완료되지 않도록)
  const setPenaltyDate = async (weekKey, memberId, performedDate) => {
    const existing = penaltyCompletions.find((p) => p.session_id === weekKey && p.member_id === memberId);
    if (existing) await updateRow('penalty_completions', 'id', existing.id, { performed_date: performedDate });
    else await insertRow('penalty_completions', { id: uid('p'), session_id: weekKey, member_id: memberId, completed_at: new Date().toISOString(), performed_date: performedDate, confirmed: false });
    await reload();
  };
  // 완료 확정/취소 — 본인 또는 운영진이 명시적으로 눌러야만 바뀜
  const setPenaltyConfirmed = async (weekKey, memberId, confirmed) => {
    const existing = penaltyCompletions.find((p) => p.session_id === weekKey && p.member_id === memberId);
    if (existing) await updateRow('penalty_completions', 'id', existing.id, { confirmed });
    else await insertRow('penalty_completions', { id: uid('p'), session_id: weekKey, member_id: memberId, completed_at: new Date().toISOString(), performed_date: todayStr(), confirmed });
    await reload();
  };
  const clearPenaltyDate = async (weekKey, memberId) => {
    const existing = penaltyCompletions.find((p) => p.session_id === weekKey && p.member_id === memberId);
    if (existing) { await deleteRow('penalty_completions', 'id', existing.id); await reload(); }
  };
  const ms = monthStr(cursor);
  const sessionsInMonth = sessions.filter((s) => s.date.startsWith(ms) && s.date <= todayStr());
  const shift = (delta) => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + delta, 1));

  // 이번 달 출석률 도트 시퀀스: 실제 세션(독서일·토론회) + 휴무일 날짜를 날짜순으로 합쳐서 표시 (휴무일도 빠짐없이 보이도록)
  const holidayDatesInMonth = calendarDays
    .filter((d) => d.date.startsWith(ms) && d.date <= todayStr() && d.type === '휴무일' && !sessions.some((s) => s.date === d.date))
    .map((d) => d.date);
  const coreDayList = [
    ...sessionsInMonth.map((s) => ({ date: s.date, session: s, inCurrentMonth: true })),
    ...holidayDatesInMonth.map((date) => ({ date, session: null, inCurrentMonth: true })),
  ].sort((a, b) => a.date.localeCompare(b.date));
  // 이번 달 첫 주가 이전 달에서 시작하면(예: 8/31 월요일 + 9/1~9/2 화·수), 월요일이 안 보여서 주가 끊겨 보이므로
  // 이전 달의 그 며칠을 시각적으로만 끌어와 채움 — 출석률 집계(denom/present)에는 포함하지 않음
  let prevMonthExtra = [];
  if (coreDayList.length > 0) {
    const mon = getMonday(coreDayList[0].date);
    const monStr = `${mon.getFullYear()}-${pad(mon.getMonth() + 1)}-${pad(mon.getDate())}`;
    if (monStr < `${ms}-01`) {
      const extraSessions = sessions.filter((s) => s.date >= monStr && s.date < `${ms}-01` && s.date <= todayStr());
      const extraHolidays = calendarDays.filter((d) => d.date >= monStr && d.date < `${ms}-01` && d.date <= todayStr() && d.type === '휴무일' && !sessions.some((s) => s.date === d.date));
      prevMonthExtra = [
        ...extraSessions.map((s) => ({ date: s.date, session: s, inCurrentMonth: false })),
        ...extraHolidays.map((date) => ({ date, session: null, inCurrentMonth: false })),
      ];
    }
  }
  const monthDayList = [...prevMonthExtra, ...coreDayList].sort((a, b) => a.date.localeCompare(b.date));
  const sessionWeekKeys = monthDayList.map((d) => weekKeyOf(d.date)); // 도트를 주 단위로 묶어서 표시하기 위함
  const weekChunkRanges = []; // [[startIdx, endIdx), ...] — 한 주(보통 월~목 4일)씩 묶은 구간
  sessionWeekKeys.forEach((wk, i) => {
    if (i === 0 || wk !== sessionWeekKeys[i - 1]) weekChunkRanges.push([i, i + 1]);
    else weekChunkRanges[weekChunkRanges.length - 1][1] = i + 1;
  });
  // 각 주 그룹이 이 달의 실제 몇 번째 주(달력 기준)인지 계산 — 모임이 달 중간부터 시작했으면 1주부터가 아니라 실제 주차부터 표기
  const sundayOf = (dateStr) => { const d = new Date(`${dateStr}T00:00:00`); d.setDate(d.getDate() - d.getDay()); return d; };
  const firstOfMonthDate = new Date(`${ms}-01T00:00:00`);
  const firstFullWeekSunday = firstOfMonthDate.getDay() === 0 ? firstOfMonthDate : new Date(firstOfMonthDate.getFullYear(), firstOfMonthDate.getMonth(), firstOfMonthDate.getDate() + (7 - firstOfMonthDate.getDay()));
  const weekOfMonth = (dateStr) => Math.max(1, Math.floor(Math.round((sundayOf(dateStr) - firstFullWeekSunday) / 86400000) / 7) + 1);
  // 주차 헤더 라벨("3주(8.17-20)")과, 그 아래 도트 행이 세로로 정확히 정렬되도록 두 행이 공유할 컬럼 폭을 미리 계산
  const weekLabels = weekChunkRanges.map(([start, end]) => {
    const chunk = monthDayList.slice(start, end);
    const d1 = chunk[0].date, d2 = chunk[chunk.length - 1].date;
    const month1 = parseInt(d1.slice(5, 7), 10), day1 = parseInt(d1.slice(8, 10), 10);
    const month2 = parseInt(d2.slice(5, 7), 10), day2 = parseInt(d2.slice(8, 10), 10);
    const range = month1 !== month2 ? `${month1}.${day1}-${month2}.${day2}` : (day1 === day2 ? `${month1}.${day1}` : `${month1}.${day1}-${day2}`);
    const repDate = chunk.find((d) => d.inCurrentMonth)?.date || d1; // 주차 번호는 실제 이번 달에 속한 날짜 기준으로 계산
    return { main: `${weekOfMonth(repDate)}주차`, sub: `(${range})` };
  });
  const weekColWidths = weekLabels.map(({ main, sub }) => Math.max(32, Math.max(main.length, sub.length) * 6.6 + 6));
  // "이번 달" 뷰는 벌칙과 동일하게 주 단위 기준으로 통일 — 월 경계에 걸쳐 끌어온 날짜(예: 8/31)도 포함해서 계산
  const totalDays = monthDayList.filter((d) => d.session).length;
  // 30분 이상: 정상 출석(1일), 15분 이상 30분 미만: 절반 인정(0.5일), 출장/휴가 사유: 별도 표시, 그 외: 결석
  const attendanceStatus = (dur) => (dur !== null && dur >= 30 ? 'full' : dur !== null && dur >= 15 ? 'half' : 'none');
  const rowsUnranked = members.map((m) => {
    const flags = monthDayList.map(({ date, session }) => {
      if (!session) return 'holiday'; // 휴무일
      const excuse = absenceExcuses.find((e) => e.date === date && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason));
      if (excuse) return excuse.reason === '휴가' ? 'vacation' : 'trip'; // 출장/휴가만 구분 (업무는 더 이상 제외 대상이 아니라 여기 걸리지 않음)
      const c = checkins.find((ck) => ck.session_id === session.id && ck.member_id === m.id);
      const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
      return attendanceStatus(dur);
    });
    let present = 0; let excusedCount = 0;
    flags.forEach((f) => {
      if (f === 'full') present += 1;
      else if (f === 'half') present += 0.5;
      else if (f === 'trip' || f === 'vacation') excusedCount += 1;
    });
    const denom = totalDays - excusedCount;
    return { ...m, present, flags, excusedCount, denom, rate: denom > 0 ? Math.round((present / denom) * 100) : 0 };
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
    let present = 0; let excusedCount = 0;
    sessions.forEach((s) => {
      const excused = absenceExcuses.some((e) => e.date === s.date && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason));
      if (excused) { excusedCount += 1; return; }
      const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id); const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null; const st = attendanceStatus(dur); present += st === 'full' ? 1 : st === 'half' ? 0.5 : 0;
    });
    const denom = totalSessions - excusedCount;
    return { ...m, present, excusedCount, denom, rate: denom > 0 ? Math.round((present / denom) * 100) : 0 };
  }).sort((a, b) => b.rate - a.rate), 'rate');

  // 월별 출석률 추이 (최근 6개월) — 클럽 전체뿐 아니라 멤버별로도 달마다 어떻게 변했는지 보여줌
  const attTrendMonths = Array.from({ length: 6 }).map((_, i) => { const d = new Date(cursor.getFullYear(), cursor.getMonth() - 5 + i, 1); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`; });
  const attTrendStats = attTrendMonths.map((mk) => {
    const monthSessions = sessions.filter((s) => s.date.startsWith(mk) && s.date <= todayStr());
    const totalDaysM = monthSessions.length;
    let totalPresent = 0; let totalDenom = 0;
    const perMember = {};
    members.forEach((m) => {
      let present = 0; let excusedCount = 0;
      monthSessions.forEach((s) => {
        const excused = absenceExcuses.some((e) => e.date === s.date && e.member_id === m.id && EXEMPT_EXCUSE_REASONS.includes(e.reason));
        if (excused) { excusedCount += 1; return; }
        const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id);
        const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
        const st = attendanceStatus(dur);
        present += st === 'full' ? 1 : st === 'half' ? 0.5 : 0;
      });
      const denom = totalDaysM - excusedCount;
      perMember[m.id] = denom > 0 ? Math.round((present / denom) * 100) : null; // null = 해당 월에 데이터 없음(휴가로 다 빠졌거나 세션 자체가 없음)
      if (denom > 0) { totalPresent += present; totalDenom += denom; }
    });
    return { mk, rate: totalDenom > 0 ? Math.round((totalPresent / totalDenom) * 100) : 0, hasData: totalDaysM > 0, perMember };
  });

  const weeklyPenalties = computeWeeklyPenalties(sessions, checkins, calendarDays, members, absenceExcuses);
  const isWeekCompleted = (wk, memberId) => penaltyCompletions.some((p) => p.session_id === wk && p.member_id === memberId && p.confirmed);
  const penaltyByMember = {};
  members.forEach((m) => { penaltyByMember[m.id] = { pending: 0 }; });
  weeklyPenalties.forEach((w) => w.results.forEach((r) => {
    if (r.missedAll && !isWeekCompleted(w.weekKey, r.member.id)) penaltyByMember[r.member.id].pending += 1;
  }));
  // 벌칙 규정 카드에 표로 보여줄 목록 — 이름 / 기준 주차 / 수행일자 / 완료여부 (최근 주가 위로 오도록 정렬)
  const penaltyEntries = [];
  weeklyPenalties.forEach((w) => w.results.forEach((r) => {
    if (!r.missedAll) return;
    const completion = penaltyCompletions.find((p) => p.session_id === w.weekKey && p.member_id === r.member.id);
    penaltyEntries.push({ member: r.member, weekKey: w.weekKey, weekLabel: monthWeekLabelOf(w.weekKey), completion });
  }));
  penaltyEntries.sort((a, b) => b.weekKey.localeCompare(a.weekKey));
  const penaltyEntriesRecent = penaltyEntries.slice(0, 10); // 대시보드엔 최근 것만 — 전체 이력은 설정 탭에서 확인

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
              <button onClick={() => shift(-1)} className="p-1.5" style={{ color: MUTE }} aria-label="이전 달"><ChevronLeft size={18} /></button>
              <div className="flex items-center gap-1.5 font-semibold" style={{ color: INK }}>
                <BookOpen size={15} style={{ color: '#F0A87C' }} />{cursor.getFullYear()}년 {cursor.getMonth() + 1}월
              </div>
              <button onClick={() => shift(1)} className="p-1.5" style={{ color: MUTE }} aria-label="다음 달"><ChevronRight size={18} /></button>
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
                const hasBusinessTrip = absenceExcuses.some((e) => e.date === date && e.reason === '출장');
                const hasVacation = absenceExcuses.some((e) => e.date === date && e.reason === '휴가');
                const hasWork = absenceExcuses.some((e) => e.date === date && e.reason === '업무');
                const hasPersonal = absenceExcuses.some((e) => e.date === date && e.reason === '개인일정');
                const hasDiscussion = types.includes('토론회');
                const myPenaltyOnDate = currentMember ? penaltyEntries.find((e) => e.member.id === currentMember.id && e.completion?.performed_date === date) : null;
                const birthdayFolks = members.filter((m) => m.birthday && mdOf(m.birthday) === date.slice(5, 10));
                const hasBirthday = birthdayFolks.length > 0;
                let borderStyle = isToday ? '1.5px solid rgba(242,238,227,0.55)' : selectedDate === date ? `1.5px solid ${textColor}` : '1px solid transparent';
                if (hasFeast) borderStyle = '1.5px solid rgba(229, 72, 77, 0.65)';
                return (
                  <button key={date} onClick={() => setSelectedDate(date === selectedDate ? null : date)}
                    className="relative aspect-square rounded-lg flex flex-col items-center justify-center text-xs leading-none"
                    style={{ background: bgStyle, color: textColor, border: borderStyle }}>
                    {(hasDiscussion || hasBirthday || myPenaltyOnDate) && (
                      <span className="absolute top-0.5 flex items-center gap-0.5">
                        {hasDiscussion && <BookOpen size={8} style={{ color: '#D9C24C' }} />}
                        {myPenaltyOnDate && <Coffee size={8} style={{ color: '#EFC94C', opacity: myPenaltyOnDate.completion?.confirmed ? 1 : 0.55 }} />}
                        {hasBirthday && (
                          <svg width="9" height="9" viewBox="0 0 24 24" fill="none">
                            <circle cx="7.5" cy="5.2" r="1.3" fill="#F0A87C" />
                            <circle cx="12" cy="4.3" r="1.3" fill="#F0A87C" />
                            <circle cx="16.5" cy="5.2" r="1.3" fill="#F0A87C" />
                            <line x1="7.5" y1="6.5" x2="7.5" y2="11" stroke="#C9A97E" strokeWidth="1.4" strokeLinecap="round" />
                            <line x1="12" y1="5.6" x2="12" y2="11" stroke="#C9A97E" strokeWidth="1.4" strokeLinecap="round" />
                            <line x1="16.5" y1="6.5" x2="16.5" y2="11" stroke="#C9A97E" strokeWidth="1.4" strokeLinecap="round" />
                            <path d="M4 21v-7a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v7z" fill="#E8927C" stroke="#B5453A" strokeWidth="1.3" strokeLinejoin="round" />
                            <path d="M4 21h16" stroke="#F2EEE3" strokeWidth="1.4" strokeLinecap="round" />
                          </svg>
                        )}
                      </span>
                    )}
                    <span>{day}</span>
                    {(hasBusinessTrip || hasVacation || hasWork || hasPersonal) && (
                      <span className="absolute bottom-0.5 flex items-center gap-0.5">
                        {hasBusinessTrip && <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#F0A87C' }}><path d="M2 15V10L6 6H21V15Z" /><rect x="9" y="8" width="4" height="3.5" /><rect x="15" y="8" width="4" height="3.5" /><line x1="2" y1="15" x2="21" y2="15" /></svg>}
                        {hasVacation && <Plane size={8} style={{ color: INK }} />}
                        {hasWork && <Briefcase size={8} style={{ color: '#D9A93A' }} />}
                        {hasPersonal && <User size={8} style={{ color: '#7FDCCF' }} />}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
            <div className="flex flex-wrap gap-2 mt-3">
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: dayTypeMeta('독서일').color }} /> 독서일</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: dayTypeMeta('토론회').color }} /> 토론회</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: 'transparent', border: '1.5px solid rgba(229, 72, 77, 0.65)' }} /> 회식일</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: dayTypeMeta('휴무일').color }} /> 휴무일</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: WEEKEND_TEXT }} /> 금·토·일(제외)</span>
            </div>
            <div className="flex flex-wrap gap-2 mt-1.5">
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}>
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none">
                  <circle cx="7.5" cy="5.2" r="1.3" fill="#F0A87C" />
                  <circle cx="12" cy="4.3" r="1.3" fill="#F0A87C" />
                  <circle cx="16.5" cy="5.2" r="1.3" fill="#F0A87C" />
                  <line x1="7.5" y1="6.5" x2="7.5" y2="11" stroke="#C9A97E" strokeWidth="1.4" strokeLinecap="round" />
                  <line x1="12" y1="5.6" x2="12" y2="11" stroke="#C9A97E" strokeWidth="1.4" strokeLinecap="round" />
                  <line x1="16.5" y1="6.5" x2="16.5" y2="11" stroke="#C9A97E" strokeWidth="1.4" strokeLinecap="round" />
                  <path d="M4 21v-7a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v7z" fill="#E8927C" stroke="#B5453A" strokeWidth="1.3" strokeLinejoin="round" />
                  <path d="M4 21h16" stroke="#F2EEE3" strokeWidth="1.4" strokeLinecap="round" />
                </svg>
                생일
              </span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#F0A87C' }}><path d="M2 15V10L6 6H21V15Z" /><rect x="9" y="8" width="4" height="3.5" /><rect x="15" y="8" width="4" height="3.5" /><line x1="2" y1="15" x2="21" y2="15" /></svg> 출장</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><Plane size={11} /> 휴가</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><Briefcase size={11} style={{ color: '#D9A93A' }} /> 업무</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><Coffee size={11} style={{ color: '#EFC94C' }} /> 내 벌칙 수행일</span>
              <span className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><User size={11} style={{ color: '#7FDCCF' }} /> 개인일정</span>
            </div>
            {(() => {
              const todayExcuses = absenceExcuses.filter((e) => e.date === todayStr());
              if (todayExcuses.length === 0) return null;
              const reasonOrder = ['출장', '휴가', '업무', '개인일정'];
              const groups = reasonOrder.map((reason) => ({
                reason,
                names: todayExcuses.filter((e) => e.reason === reason).map((e) => members.find((m) => m.id === e.member_id)).filter(Boolean).map((m) => dispName(m.name, isLoggedIn)),
              })).filter((g) => g.names.length > 0);
              if (groups.length === 0) return null;
              return (
                <div className="mt-3 p-3 rounded-xl" style={{ background: 'rgba(240,168,124,0.09)', border: '1px solid rgba(240,168,124,0.32)' }}>
                  <div className="text-xs mb-1.5" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>오늘 참석불가</div>
                  <div className="space-y-1">
                    {groups.map((g) => (
                      <div key={g.reason} className="text-xs" style={{ color: NEUTRAL_TEXT }}>
                        <span style={{ color: MUTE }}>{g.reason} : </span>{g.names.join(', ')}
                      </div>
                    ))}
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
                          <button onClick={() => setMyExcuseForDate('업무')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>업무</button>
                          <button onClick={() => setMyExcuseForDate('개인일정')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>개인일정</button>
                        </div>
                      )}
                    </div>
                  )}
                  {currentMember && penaltyEntries.some((e) => e.member.id === currentMember.id) && (
                    <div className="mb-3">
                      <div className="text-xs mb-1.5 flex items-center gap-1" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}><Coffee size={12} style={{ color: '#EFC94C' }} /> 본인 벌칙 수행일 지정</div>
                      <div className="space-y-1.5">
                        {penaltyEntries.filter((e) => e.member.id === currentMember.id).map((e) => {
                          const isSetToThisDate = e.completion?.performed_date === selectedDate;
                          return (
                            <div key={e.weekKey} className="flex items-center justify-between text-xs">
                              <span style={{ color: NEUTRAL_TEXT }}>{e.weekLabel}</span>
                              {isSetToThisDate ? (
                                <button onClick={() => clearPenaltyDate(e.weekKey, currentMember.id)} className="text-[11px] underline underline-offset-2" style={{ color: MUTE }}>이 날짜 지정 취소</button>
                              ) : (
                                <button onClick={() => setPenaltyDate(e.weekKey, currentMember.id, selectedDate)} className="inline-flex items-center gap-1 rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}><Coffee size={11} /> 이 날짜로 지정(예정)</button>
                              )}
                            </div>
                          );
                        })}
                      </div>
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

          <Card>
            <div className="text-sm font-semibold mb-2" style={{ color: INK }}>이번 달 출석률</div>
            <div className="flex flex-wrap items-center gap-2.5 mb-3">
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, background: '#7FA8D9' }} />출석</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, background: 'linear-gradient(90deg, #7FA8D9 50%, transparent 50%)', border: `1px solid ${LINE}` }} />절반출석(15~29분)</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, background: '#E0958C' }} />휴무일</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#F0A87C' }}><path d="M2 15V10L6 6H21V15Z" /><rect x="9" y="8" width="4" height="3.5" /><rect x="15" y="8" width="4" height="3.5" /><line x1="2" y1="15" x2="21" y2="15" /></svg>출장</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><Plane size={9} style={{ color: INK }} />휴가</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, border: `1px solid ${LINE}` }} />결석</span>
            </div>
            {weekChunkRanges.length > 0 && (
              <div className="flex items-center flex-wrap gap-y-1 mb-1.5" style={{ paddingLeft: 34 }}>
                {weekChunkRanges.map(([start, end], wi) => (
                  <React.Fragment key={wi}>
                    <div className="flex flex-col items-center" style={{ width: weekColWidths[wi] }}>
                      <span className="text-[10px] leading-tight" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: 'nowrap' }}>{weekLabels[wi].main}</span>
                      <span className="text-[10px] leading-tight" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: 'nowrap' }}>{weekLabels[wi].sub}</span>
                    </div>
                    {wi < weekChunkRanges.length - 1 && (
                      <span className="shrink-0" style={{ width: 1, height: 11, background: MUTE, opacity: 0.4, transform: 'rotate(22deg)', margin: '0 7px' }} />
                    )}
                  </React.Fragment>
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
                      {isCurrentMonth && warningMemberIds.has(r.id) && !(penaltyByMember[r.id]?.pending > 0) && <span title="이번 주 열린 독서일을 지금까지 모두 결석 — 벌칙유의" className="inline-flex items-center gap-0.5 rounded-full px-1.5 py-0.5 text-[10px] font-semibold shrink-0" style={{ background: '#3A2213', color: '#F0A87C' }}>⚠️ 벌칙유의</span>}
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      {penaltyByMember[r.id]?.pending > 0 && <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}><Gavel size={10} /> 벌칙 대상 {penaltyByMember[r.id].pending}</span>}
                      <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.present}/{r.denom} · {r.rate}%</span>
                    </div>
                  </div>
                  <div className="flex items-center flex-wrap gap-y-1.5" style={{ paddingLeft: 34 }}>
                    {weekChunkRanges.map(([start, end], wi) => (
                      <React.Fragment key={wi}>
                        <div className="flex items-center justify-center gap-1" style={{ width: weekColWidths[wi] }}>
                          {r.flags.slice(start, end).map((status, i) => {
                            const fromPrevMonth = !monthDayList[start + i].inCurrentMonth;
                            if (status === 'trip') {
                              return <svg key={i} width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#F0A87C', opacity: fromPrevMonth ? 0.6 : 1 }}><path d="M2 15V10L6 6H21V15Z" /><rect x="9" y="8" width="4" height="3.5" /><rect x="15" y="8" width="4" height="3.5" /><line x1="2" y1="15" x2="21" y2="15" /></svg>;
                            }
                            if (status === 'vacation') {
                              return <Plane key={i} size={9} style={{ color: INK, opacity: fromPrevMonth ? 0.6 : 1 }} />;
                            }
                            return (
                              <span key={i} className="relative rounded-full" style={{
                                width: 9, height: 9,
                                background: status === 'full' ? '#7FA8D9' : status === 'half' ? 'linear-gradient(90deg, #7FA8D9 50%, transparent 50%)' : status === 'holiday' ? '#E0958C' : 'transparent',
                                border: status === 'full' || status === 'holiday' ? 'none' : `1.5px solid ${LINE}`,
                              }} title={(status === 'holiday' ? '휴무일' : '') + (fromPrevMonth ? ' (이전 달, 참고용)' : '')}>
                                {fromPrevMonth && <span className="absolute inset-0 rounded-full pointer-events-none" style={{ background: 'repeating-linear-gradient(45deg, rgba(30,28,22,0.55) 0px, rgba(30,28,22,0.55) 1px, transparent 1px, transparent 2.5px)' }} />}
                              </span>
                            );
                          })}
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
            <p className="text-[10px] mt-3 pt-2" style={{ color: MUTE, borderTop: `1px solid ${ROW_LINE}` }}>※ 출석률 산정제외 : 출장, 휴가</p>
            <p className="text-[10px] mt-1" style={{ color: MUTE }}>※ 출석률 산정기준 : 주 단위(월 경계에 걸친 경우, 전월분 포함 산정)</p>
          </Card>
          {penaltyRule && (penaltyEntries.length > 0 || warningMemberIds.size > 0) && (
            <Card>
              <div className="flex items-center gap-1.5 text-sm font-semibold mb-1" style={{ color: INK }}><Gavel size={16} style={{ color: '#F0A87C' }} /> 벌칙 현황</div>
              <p className="text-sm whitespace-pre-wrap" style={{ color: NEUTRAL_TEXT }}>{penaltyRule}</p>
              {penaltyEntriesRecent.length > 0 && (
                <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs" style={{ color: MUTE }}>벌칙 대상자 (최근 10건)</span>
                    <span className="text-[10px]" style={{ color: MUTE }}>완료 여부</span>
                  </div>
                  <div className="space-y-2">
                    {penaltyEntriesRecent.map((e, i) => {
                      const canEdit = e.member.id === currentMember?.id || canManage;
                      const rowKey = `${e.weekKey}_${e.member.id}`;
                      const isExpanded = expandedPenaltyRow === rowKey;
                      const confirmed = !!e.completion?.confirmed;
                      const hasDate = !!e.completion?.performed_date;
                      const draft = penaltyDateDrafts[rowKey] ?? e.completion?.performed_date ?? todayStr();
                      return (
                        <div key={rowKey} className="text-xs">
                          <div className="flex items-center justify-between gap-2 flex-wrap">
                            <div className="flex items-center gap-2 min-w-0">
                              <Stamp role={e.member.role} size={20} tilt={0} />
                              <span className="truncate" style={{ color: INK }}>{dispName(e.member.name, isLoggedIn)}</span>
                              <span style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{e.weekLabel}</span>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0">
                              {confirmed ? (
                                <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 font-semibold" style={{ background: '#12302C', color: '#7FDCCF' }}><Check size={10} /> 완료 · {fmtDate(e.completion.performed_date)}</span>
                              ) : hasDate ? (
                                <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 font-semibold" style={{ background: '#3A2E10', color: '#EFC94C' }}>예정 · {fmtDate(e.completion.performed_date)}</span>
                              ) : (
                                <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}>미완료</span>
                              )}
                              {canEdit && (
                                <button onClick={() => setExpandedPenaltyRow(isExpanded ? null : rowKey)} className="p-1" style={{ color: MUTE }} aria-label="벌칙 정보 수정">
                                  <Pencil size={12} />
                                </button>
                              )}
                            </div>
                          </div>
                          {isExpanded && canEdit && (
                            <div className="flex items-center flex-wrap gap-1.5 mt-1.5 pl-6">
                              <input type="date" value={draft} onChange={(ev) => setPenaltyDateDrafts((prev) => ({ ...prev, [rowKey]: ev.target.value }))}
                                className="rounded-lg border px-1.5 py-1 text-[11px] outline-none" style={inputStyle} aria-label="벌칙 수행 예정일" />
                              <button onClick={() => setPenaltyDate(e.weekKey, e.member.id, draft)} className="rounded-full px-2 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>예정일 저장</button>
                              {confirmed ? (
                                <button onClick={() => setPenaltyConfirmed(e.weekKey, e.member.id, false)} className="rounded-full px-2 py-1 font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}>완료 취소</button>
                              ) : (
                                <button onClick={() => setPenaltyConfirmed(e.weekKey, e.member.id, true)} className="inline-flex items-center gap-1 rounded-full px-2 py-1 font-semibold" style={{ background: '#12302C', color: '#7FDCCF' }}><Check size={10} /> 완료로 확정</button>
                              )}
                              {hasDate && <button onClick={() => clearPenaltyDate(e.weekKey, e.member.id)} className="text-[11px] underline underline-offset-2" style={{ color: MUTE }}>일정 삭제</button>}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                  <p className="text-[10px] mt-2" style={{ color: MUTE }}>날짜만 저장해선 자동으로 완료되지 않아요 — 본인 또는 운영진이 "완료로 확정"을 눌러야 완료 처리돼요. 연필 아이콘을 눌러 수정할 수 있어요.</p>
                </div>
              )}
            </Card>
          )}
        </>
      ) : (
        <>
          <Card>
            <div className="text-sm font-semibold mb-3" style={{ color: INK }}>월별 출석률 추이 (최근 6개월)</div>
            <div className="flex items-end justify-between gap-1.5" style={{ height: 88 }}>
              {attTrendStats.map((t) => (
                <div key={t.mk} className="flex-1 flex items-end justify-center" style={{ height: '100%' }}>
                  <div className="rounded-t-sm" style={{ width: 16, height: `${Math.max(2, t.rate)}%`, background: t.hasData ? (t.rate >= 80 ? '#7FDCCF' : t.rate >= 50 ? '#EFC94C' : '#F0A87C') : 'transparent', border: t.hasData ? 'none' : `1px dashed ${LINE}` }} title={t.hasData ? `${t.rate}%` : '데이터 없음'} />
                </div>
              ))}
            </div>
            <div className="flex items-center justify-between mt-1.5">
              {attTrendStats.map((t) => (
                <span key={t.mk} className="flex-1 text-center text-[10px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{t.mk.slice(5)}월</span>
              ))}
            </div>
            <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
              <div className="text-xs mb-2" style={{ color: MUTE }}>멤버별 월별 출석률</div>
              <div className="overflow-x-auto">
                <table className="w-full text-center" style={{ borderCollapse: 'collapse' }}>
                  <thead>
                    <tr>
                      <th className="text-left text-[10px] pb-1.5" style={{ color: MUTE, fontWeight: 400 }}></th>
                      {attTrendStats.map((t) => (
                        <th key={t.mk} className="text-[10px] pb-1.5 px-1" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 400 }}>{t.mk.slice(5)}월</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {members.map((m) => (
                      <tr key={m.id}>
                        <td className="text-left text-xs py-1 pr-2 whitespace-nowrap" style={{ color: INK }}>{dispName(m.name, isLoggedIn)}</td>
                        {attTrendStats.map((t) => {
                          const rate = t.perMember[m.id];
                          const color = rate === null ? MUTE : rate >= 80 ? '#7FDCCF' : rate >= 50 ? '#EFC94C' : '#F0A87C';
                          return (
                            <td key={t.mk} className="text-[11px] py-1 px-1" style={{ color, fontFamily: "'IBM Plex Mono', monospace" }}>{rate === null ? '–' : `${rate}%`}</td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <p className="text-[10px] mt-3 pt-2" style={{ color: MUTE, borderTop: `1px solid ${ROW_LINE}` }}>※ 산정기준 : 월 단위</p>
          </Card>
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
                        <span style={{ fontSize: 10, color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.present}/{r.denom}</span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
            <p className="text-[10px] mt-3 pt-2" style={{ color: MUTE, borderTop: `1px solid ${ROW_LINE}` }}>※ 출석률 산정제외 : 출장, 휴가</p>
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
  const [editName, setEditName] = useState(''); const [editRole, setEditRole] = useState('회원'); const [editBirthday, setEditBirthday] = useState(''); const [editPin, setEditPin] = useState(''); const [clearPin, setClearPin] = useState(false);
  const [editDept, setEditDept] = useState(''); const [editJobType, setEditJobType] = useState(''); const [editJoinedAt, setEditJoinedAt] = useState(''); const [editGenre, setEditGenre] = useState(''); const [editNote, setEditNote] = useState('');
  // 멤버 탭 화면 표시용 — 로그인한 본인이 맨 위로 오도록 재배치 (다른 탭에서 쓰는 sortedMembers 순서엔 영향 없음)
  const displayMembers = isLoggedIn ? [...sortedMembers].sort((a, b) => (a.id === currentUserId ? -1 : b.id === currentUserId ? 1 : 0)) : sortedMembers;

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
  // PIN은 목록 조회에 안 실려 있어서(m.has_pin만 boolean으로 존재), 수정 시작 시 실제 값은 프리필하지 않음 — 빈 칸=기존 값 유지, 입력하면 새 값으로 교체
  const startEdit = (m) => { setEditingId(m.id); setEditMode('full'); setEditName(m.name); setEditRole(m.role); setEditBirthday(m.birthday || ''); setEditPin(''); setClearPin(false); setEditDept(m.department || ''); setEditJobType(m.job_type || ''); setEditJoinedAt(m.joined_at || ''); setEditGenre(m.book_genre || ''); setEditNote(m.note || ''); };
  const startSelfEdit = (m) => { setEditingId(m.id); setEditMode('self'); setEditBirthday(m.birthday || ''); setEditPin(''); setClearPin(false); setEditDept(m.department || ''); setEditJobType(m.job_type || ''); setEditJoinedAt(m.joined_at || ''); setEditGenre(m.book_genre || ''); setEditNote(m.note || ''); };
  const saveEdit = async () => {
    if (editPin && !/^\d{4}$/.test(editPin)) return;
    const pinPatch = clearPin ? { pin: null } : editPin ? { pin: editPin } : {}; // 빈 칸이면 pin 필드 자체를 patch에서 빼서 기존 값 그대로 유지
    if (editMode === 'full') {
      if (!editName.trim()) return;
      await updateRow('members', 'id', editingId, { name: editName.trim(), role: editRole, birthday: editBirthday || null, department: editDept || null, job_type: editJobType || null, joined_at: editJoinedAt || null, book_genre: editGenre || null, note: editNote || null, ...pinPatch });
    } else {
      await updateRow('members', 'id', editingId, { birthday: editBirthday || null, department: editDept || null, job_type: editJobType || null, book_genre: editGenre || null, ...pinPatch });
    }
    await reload();
    setEditingId(null);
  };
  const removeMember = async (id) => { await deleteRow('members', 'id', id); await reload(); if (currentUserId === id) setIdentity(null); };

  const downloadExcel = async () => {
    // PIN은 목록 상태에 없으므로, 다운로드하는 이 순간에만 실제 members 테이블에서 좁게 조회 (관리자가 명시적으로 요청했을 때만 전송됨)
    const { data: withPins } = await supabase.from('members').select('id,pin');
    const pinById = Object.fromEntries((withPins || []).map((r) => [r.id, r.pin]));
    const data = sortedMembers.map((m) => ({ 이름: m.name, 직급: m.role, 소속: m.department || '', 직군: m.job_type || '', 생일: m.birthday || '', 가입일자: m.joined_at || '', 선호도서: m.book_genre || '', 비고: m.note || '', PIN: pinById[m.id] || '' }));
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
        {displayMembers.length === 0 && <div className="px-4 py-8 text-center text-sm" style={{ color: MUTE }}>등록된 멤버가 없어요.</div>}
        {displayMembers.map((m, idx) => (
          <div key={m.id} style={{ borderTop: idx === 0 ? 'none' : `1px solid ${ROW_LINE}` }}>
            {editingId === m.id ? (
              <div className="p-4 space-y-3" style={{ background: '#1A1812' }}>
                <div className="flex items-center gap-2 rounded-xl px-3 py-2" style={{ background: '#3A2E10' }}>
                  <Stamp role={m.role} size={28} tilt={0} />
                  <span className="font-semibold" style={{ color: '#EFC94C' }}>{m.name} 정보 수정</span>
                </div>
                {editMode === 'full' && (
                  <>
                    <input value={editName} onChange={(e) => setEditName(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="이름" />
                    <RolePicker value={editRole} onChange={setEditRole} />
                  </>
                )}
                {editMode === 'self' && <p className="text-xs" style={{ color: MUTE }}>이름·직급·가입일자·비고는 회장·간사·총무만 변경할 수 있어요. 나머지 정보는 본인이 직접 수정할 수 있어요.</p>}
                <div className="grid grid-cols-2 gap-2">
                  <input value={editDept} onChange={(e) => setEditDept(e.target.value)} placeholder="소속" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="소속" />
                  <input value={editJobType} onChange={(e) => setEditJobType(e.target.value)} placeholder="직군" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="직군" />
                </div>
                {editMode === 'full' && (
                  <>
                    <div><label htmlFor={`edit-joined-${m.id}`} className="text-xs mb-1 block" style={{ color: MUTE }}>가입일자</label><input id={`edit-joined-${m.id}`} type="date" value={editJoinedAt} onChange={(e) => setEditJoinedAt(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
                  </>
                )}
                <input value={editGenre} onChange={(e) => setEditGenre(e.target.value)} placeholder="선호 도서 종류 (예: 소설, 자기계발)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="선호 도서 종류" />
                {editMode === 'full' && (
                  <input value={editNote} onChange={(e) => setEditNote(e.target.value)} placeholder="비고" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="비고" />
                )}
                <div><label htmlFor={`edit-birthday-${m.id}`} className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Cake size={13} /> 생일</label><input id={`edit-birthday-${m.id}`} type="date" value={editBirthday} onChange={(e) => setEditBirthday(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
                <div>
                  <label htmlFor={`edit-pin-${m.id}`} className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Lock size={13} /> 본인 확인 PIN (4자리, 선택)</label>
                  <input id={`edit-pin-${m.id}`} type="password" inputMode="numeric" maxLength={4} value={editPin} disabled={clearPin}
                    onChange={(e) => setEditPin(e.target.value.replace(/\D/g, ''))}
                    className="w-full rounded-xl border px-3 py-2 text-sm tracking-[0.3em] outline-none" style={{ ...inputStyle, opacity: clearPin ? 0.5 : 1 }}
                    placeholder={m.has_pin ? '변경하려면 입력 (비워두면 기존 PIN 유지)' : '설정 안 함'} />
                  {m.has_pin && (
                    <label className="inline-flex items-center gap-1.5 mt-1.5 text-xs" style={{ color: MUTE }}>
                      <input type="checkbox" checked={clearPin} onChange={(e) => { setClearPin(e.target.checked); if (e.target.checked) setEditPin(''); }} />
                      기존 PIN 삭제(다음부턴 PIN 없이 로그인)
                    </label>
                  )}
                </div>
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
                      {m.has_pin && <Lock size={11} style={{ color: MUTE }} />}
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
                  {canManage && <button onClick={() => startEdit(m)} className="p-2 rounded-lg" style={{ color: MUTE }} aria-label="회원 정보 수정"><Pencil size={16} /></button>}
                  {!canManage && m.id === currentUserId && <button onClick={() => startSelfEdit(m)} className="p-2 rounded-lg" style={{ color: MUTE }} aria-label="내 정보 수정"><Pencil size={16} /></button>}
                  {canManage && <button onClick={() => requestDelete(() => removeMember(m.id), `${m.name}님을 삭제할까요? 관련 기록도 함께 사라져요.`)} className="p-2 rounded-lg" style={{ color: '#F0A87C' }} aria-label="회원 삭제"><Trash2 size={16} /></button>}
                </div>
              </div>
            )}
          </div>
        ))}
      </Card>

      {(canManage || sortedMembers.length === 0) && (
        showAddForm ? (
          <Card className="space-y-3">
            <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="이름을 입력하세요" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="이름" />
            <RolePicker value={newRole} onChange={setNewRole} />
            <div className="grid grid-cols-2 gap-2">
              <input value={newDept} onChange={(e) => setNewDept(e.target.value)} placeholder="소속 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="소속" />
              <input value={newJobType} onChange={(e) => setNewJobType(e.target.value)} placeholder="직군 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="직군" />
            </div>
            <div><label htmlFor="new-joined" className="text-xs mb-1 block" style={{ color: MUTE }}>가입일자 (선택)</label><input id="new-joined" type="date" value={newJoinedAt} onChange={(e) => setNewJoinedAt(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            <input value={newGenre} onChange={(e) => setNewGenre(e.target.value)} placeholder="선호 도서 종류 (선택)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="선호 도서 종류" />
            <input value={newNote} onChange={(e) => setNewNote(e.target.value)} placeholder="비고 (선택)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="비고" />
            <div><label htmlFor="new-birthday" className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Cake size={13} /> 생일 (선택)</label><input id="new-birthday" type="date" value={newBirthday} onChange={(e) => setNewBirthday(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            <div><label htmlFor="new-pin" className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Lock size={13} /> 본인 확인 PIN (4자리, 선택)</label><input id="new-pin" type="password" inputMode="numeric" maxLength={4} value={newPin} onChange={(e) => setNewPin(e.target.value.replace(/\D/g, ''))} className="w-full rounded-xl border px-3 py-2 text-sm tracking-[0.3em] outline-none" style={inputStyle} placeholder="설정 안 함" /></div>
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
          <button onClick={() => shift(-1)} className="p-1.5" style={{ color: MUTE }} aria-label="이전 달"><ChevronLeft size={18} /></button>
          <div className="font-semibold" style={{ color: INK }}>{cursor.getFullYear()}년 {cursor.getMonth() + 1}월</div>
          <button onClick={() => shift(1)} className="p-1.5" style={{ color: MUTE }} aria-label="다음 달"><ChevronRight size={18} /></button>
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
                      <button onClick={() => requestDelete(() => removeExpense(e.id), '이 지출 내역을 삭제할까요?')} className="p-1" style={{ color: MUTE }} aria-label="지출 삭제"><Trash2 size={14} /></button>
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
          <input value={expDesc} onChange={(e) => setExpDesc(e.target.value)} placeholder="내역 (예: 간식비)" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="지출 내역" />
        </div>
        <div className="flex gap-2">
          <input type="number" value={expAmount} onChange={(e) => setExpAmount(e.target.value)} placeholder="금액" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="지출 금액" />
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
                <button onClick={() => requestDelete(() => removeExpense(e.id), '이 지출 내역을 삭제할까요?')} className="p-1" style={{ color: MUTE }} aria-label="지출 삭제"><Trash2 size={14} /></button>
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
  const [visitDetailsOpen, setVisitDetailsOpen] = useState(false);

  // 접속자수 조회 — 간사만 볼 수 있음, site_visits 테이블에서 별도 조회(전체 reload 사이클과 무관). 날짜를 골라서 조회 가능
  const isSecretary = currentMember?.role === '간사';
  const [visitQueryDate, setVisitQueryDate] = useState(todayStr());
  const [visitQueryCount, setVisitQueryCount] = useState(null);
  const [visitQueryLog, setVisitQueryLog] = useState(null); // [{name, time}, ...] 해당 날짜에 로그인 상태로 접속한 기록(시각 포함), 최신순
  const [visitHistory, setVisitHistory] = useState(null); // [{date, count}, ...] 최근 7일
  useEffect(() => {
    if (!isSecretary) return;
    const start = `${visitQueryDate}T00:00:00`;
    const end = `${visitQueryDate}T23:59:59`;
    supabase.from('site_visits').select('member_id, visited_at').gte('visited_at', start).lte('visited_at', end)
      .then(({ data }) => {
        setVisitQueryCount(data?.length ?? 0);
        const log = (data || [])
          .filter((r) => r.member_id)
          .map((r) => ({ name: members.find((m) => m.id === r.member_id)?.name, time: r.visited_at }))
          .filter((r) => r.name)
          .sort((a, b) => b.time.localeCompare(a.time));
        setVisitQueryLog(log);
      })
      .catch(() => { setVisitQueryCount(null); setVisitQueryLog(null); });
  }, [isSecretary, visitQueryDate]);
  useEffect(() => {
    if (!isSecretary) return;
    const since = new Date(); since.setDate(since.getDate() - 6); since.setHours(0, 0, 0, 0);
    supabase.from('site_visits').select('visited_at').gte('visited_at', since.toISOString())
      .then(({ data }) => {
        const counts = {};
        (data || []).forEach((row) => { const d = row.visited_at.slice(0, 10); counts[d] = (counts[d] || 0) + 1; });
        const days = Array.from({ length: 7 }).map((_, i) => {
          const d = new Date(since); d.setDate(d.getDate() + i);
          const key = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
          return { date: key, count: counts[key] || 0 };
        });
        setVisitHistory(days);
      })
      .catch(() => setVisitHistory(null));
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
  const updateCheckin = async (id, field, timeVal) => { if (!timeVal) return; await updateRow('checkins', 'id', id, { [field]: new Date(`${date}T${timeVal}`).toISOString() }); await reload(['checkins']); };
  const removeCheckin = async (id) => { await deleteRow('checkins', 'id', id); await reload(['checkins']); };
  const saveRule = () => { setPenaltyRule(ruleInput.trim()); setEditingRule(false); };
  const weeklyPenalties = computeWeeklyPenalties(sessions, checkins, calendarDays, members, absenceExcuses);
  const addExcuse = async (memberId, reason) => {
    const existing = absenceExcuses.find((e) => e.date === date && e.member_id === memberId);
    if (existing) await deleteRow('absence_excuses', 'id', existing.id);
    await insertRow('absence_excuses', { id: uid('ae'), date, member_id: memberId, reason });
    await reload();
  };
  const removeExcuse = async (id) => { await deleteRow('absence_excuses', 'id', id); await reload(); };
  const isWeekCompleted = (wk, memberId) => penaltyCompletions.some((p) => p.session_id === wk && p.member_id === memberId && p.confirmed);
  const toggleWeekCompletion = async (wk, memberId, performedDate) => {
    const existing = penaltyCompletions.find((p) => p.session_id === wk && p.member_id === memberId);
    if (existing?.confirmed) await updateRow('penalty_completions', 'id', existing.id, { confirmed: false }); // 완료 취소 (기록은 남김)
    else if (existing) await updateRow('penalty_completions', 'id', existing.id, { confirmed: true, performed_date: performedDate || existing.performed_date || todayStr() }); // 예정 상태였던 걸 확정으로 전환
    else await insertRow('penalty_completions', { id: uid('p'), session_id: wk, member_id: memberId, completed_at: new Date().toISOString(), performed_date: performedDate || todayStr(), confirmed: true }); // 완전히 새로 생성+확정
    await reload();
  };
  const [performDateInputs, setPerformDateInputs] = useState({}); // { `${wk}_${memberId}`: 'YYYY-MM-DD' } — 완료 처리 전 날짜 선택용
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
                <button onClick={() => requestDelete(() => removeCheckin(c.id), '이 출결 기록을 삭제할까요?')} className="p-1" style={{ color: '#F0A87C' }} aria-label="출결 기록 삭제"><Trash2 size={14} /></button>
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
                    <button onClick={() => removeExcuse(excuse.id)} className="p-1" style={{ color: MUTE }} aria-label="사유 삭제"><X size={14} /></button>
                  </div>
                ) : (
                  <div className="flex gap-1.5 shrink-0 flex-wrap justify-end">
                    <button onClick={() => addExcuse(m.id, '출장')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>출장</button>
                    <button onClick={() => addExcuse(m.id, '휴가')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>휴가</button>
                    <button onClick={() => addExcuse(m.id, '업무')} className="text-xs rounded-full px-2.5 py-1 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>업무</button>
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
        <div className="flex gap-2"><input type="time" value={manualIn} onChange={(e) => setManualIn(e.target.value)} className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="체크인 시각" /><input type="time" value={manualOut} onChange={(e) => setManualOut(e.target.value)} className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} aria-label="체크아웃 시각" /></div>
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
          {!editingRule && <button onClick={() => { setRuleInput(penaltyRule || ''); setEditingRule(true); }} className="p-1.5" style={{ color: MUTE }} aria-label="벌칙 규정 수정"><Pencil size={14} /></button>}
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
                    <div className="text-sm font-semibold" style={{ color: INK }}>{fmtDate(w.sessions[0])} ~ {fmtDate(w.sessions[3])}</div>
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
                      const completion = penaltyCompletions.find((p) => p.session_id === w.weekKey && p.member_id === r.member.id);
                      const inputKey = `${w.weekKey}_${r.member.id}`;
                      return (
                        <div key={r.member.id} className="flex items-center justify-between text-xs py-1 gap-2">
                          <div className="flex items-center gap-2 shrink-0"><Stamp role={r.member.role} size={22} tilt={0} /><span style={{ color: NEUTRAL_TEXT }}>{r.member.name}</span></div>
                          {done ? (
                            <div className="flex items-center gap-2">
                              <span className="text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{completion?.performed_date ? fmtDate(completion.performed_date) : '날짜 미기록'} 수행 완료</span>
                              <button onClick={() => toggleWeekCompletion(w.weekKey, r.member.id)} className="inline-flex items-center gap-1 rounded-full px-2 py-1 font-semibold shrink-0" style={{ background: '#12302C', color: '#7FDCCF' }}><Check size={11} /> 취소</button>
                            </div>
                          ) : (
                            <div className="flex items-center gap-1.5">
                              {completion?.performed_date && <span className="text-[11px]" style={{ color: '#EFC94C', fontFamily: "'IBM Plex Mono', monospace" }}>예정 · {fmtDate(completion.performed_date)}</span>}
                              <input type="date" value={performDateInputs[inputKey] || completion?.performed_date || todayStr()} onChange={(e) => setPerformDateInputs((prev) => ({ ...prev, [inputKey]: e.target.value }))}
                                className="rounded-lg border px-1.5 py-1 text-[11px] outline-none" style={inputStyle} aria-label="벌칙 수행일자" />
                              <button onClick={() => toggleWeekCompletion(w.weekKey, r.member.id, performDateInputs[inputKey])} className="rounded-full px-2 py-1 font-semibold shrink-0" style={{ background: '#3A2213', color: '#F0A87C' }}>완료로 확정</button>
                            </div>
                          )}
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
      {isSecretary && (
        <Card>
          <button onClick={() => setVisitDetailsOpen((v) => !v)} className="flex items-center gap-1.5" aria-label={visitDetailsOpen ? '접속자수 상세 접기' : '접속자수 상세 펼치기'}>
            <Settings2 size={18} style={{ color: MUTE }} />
            {visitDetailsOpen ? <ChevronUp size={14} style={{ color: MUTE }} /> : <ChevronDown size={14} style={{ color: MUTE }} />}
          </button>
          {visitDetailsOpen && (
            <>
              <div className="flex items-center justify-between mt-3 mb-2">
                <span className="text-sm font-semibold" style={{ color: INK }}>접속자수</span>
                <input type="date" value={visitQueryDate} onChange={(e) => setVisitQueryDate(e.target.value)} className="rounded-lg border px-2 py-1 text-xs outline-none" style={inputStyle} />
              </div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs" style={{ color: MUTE }}>{fmtDate(visitQueryDate)} 접속</span>
                <span className="text-lg font-semibold" style={{ color: '#7FDCCF', fontFamily: "'IBM Plex Mono', monospace" }}>{visitQueryCount === null ? '—' : `${visitQueryCount}회`}</span>
              </div>
              {visitQueryLog && (
                <div className="text-xs pt-2" style={{ color: MUTE, borderTop: `1px solid ${ROW_LINE}` }}>
                  {visitQueryLog.length > 0 ? (
                    <div className="space-y-0.5">
                      <div className="mb-1">로그인 상태로 접속:</div>
                      {visitQueryLog.map((v, i) => (
                        <div key={i} className="flex items-center justify-between">
                          <span style={{ color: NEUTRAL_TEXT }}>{v.name}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{new Date(v.time).toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', hour12: false })}</span>
                        </div>
                      ))}
                    </div>
                  ) : '이 날짜엔 로그인 상태로 접속한 멤버가 없어요.'}
                </div>
              )}
              {visitHistory && (
                <div className="space-y-1 pt-2 mt-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="text-xs mb-1" style={{ color: MUTE }}>최근 7일 이력</div>
                  {visitHistory.map((d) => (
                    <div key={d.date} className="flex items-center justify-between text-xs py-0.5">
                      <span style={{ color: NEUTRAL_TEXT }}>{fmtDate(d.date)}</span>
                      <span style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{d.count}회</span>
                    </div>
                  ))}
                </div>
              )}
              <p className="text-[11px] mt-2" style={{ color: MUTE }}>앱이 열린 횟수예요 (같은 사람이 여러 번 들어오면 중복 집계될 수 있어요). 간사에게만 보여요.</p>
            </>
          )}
        </Card>
      )}
    </div>
  );
}
