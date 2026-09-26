import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './supabaseClient';
import {
  Crown, Shield, Wallet, User, Plus, Pencil, Trash2, Check, X, Lock, AlertCircle, Mail,
  Megaphone, QrCode, BarChart3, Users, Settings2, Settings, Download, Upload, ChevronLeft, ChevronRight, ChevronUp, ChevronDown, Briefcase, Coffee,
  LogIn, LogOut, Cake, PartyPopper, Archive, Paperclip, FileText, Eye, Pin, Gavel, BookOpen, Search,
  Image as ImageIcon, Trophy, Plane, Library,
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
// 북적북적 책탑 - 가죽 책등 느낌의 색상 팔레트 (책마다 id 기반으로 하나씩 고정 배정)
const LEATHER_PALETTE = [
  { bg: 'linear-gradient(180deg, #AD7A4E 0%, #8B5E34 45%, #6F461D 100%)', text: '#FBF3E4', line: 'rgba(255,235,205,0.55)' },
  { bg: 'linear-gradient(180deg, #52807D 0%, #3E6B69 45%, #2C4E4C 100%)', text: '#F0F7F5', line: 'rgba(220,240,235,0.5)' },
  { bg: 'linear-gradient(180deg, #E6D9BD 0%, #D9C7A3 45%, #C2A87C 100%)', text: '#3A2C18', line: 'rgba(60,40,10,0.35)' },
  { bg: 'linear-gradient(180deg, #5E3D2C 0%, #472B1D 45%, #331E13 100%)', text: '#F3E7D8', line: 'rgba(255,230,200,0.4)' },
  { bg: 'linear-gradient(180deg, #446485 0%, #2E4A66 45%, #1F3650 100%)', text: '#EDF2F7', line: 'rgba(220,235,250,0.45)' },
  { bg: 'linear-gradient(180deg, #4F7350 0%, #3A5A3A 45%, #293F29 100%)', text: '#EFF5EC', line: 'rgba(220,240,215,0.4)' },
  { bg: 'linear-gradient(180deg, #4F4F4D 0%, #3A3A38 45%, #292927 100%)', text: '#F0EEE9', line: 'rgba(230,228,220,0.35)' },
  { bg: 'linear-gradient(180deg, #7A4444 0%, #663636 45%, #4A2626 100%)', text: '#F5E9E4', line: 'rgba(250,225,215,0.4)' },
];
// 도서 공유함 구분 색상 - 상태 뱃지(대여가능=민트, 요청중/대여중=골드)와 겹치지 않는 색으로 분리
const SHARE_OFFER_COLOR = '#B39DDB';   // 빌려줄까요? - 보라색
const SHARE_OFFER_BG = '#2A2438';
const SHARE_REQUEST_COLOR = '#F0A87C'; // 빌려주실 수 있나요? - 코랄
const SHARE_REQUEST_BG = '#3A2519';
// 글쓰기/추가 패널 - 아래 목록과 확실히 구분되도록 카드 안의 별도 박스로 표시 (도서공유함·북적북적 공통)
const FORM_PANEL_BG = '#171510';
const FORM_PANEL_BORDER = '#3D3826';
// 책장 상태 리본 색상
const RIBBON_DONE = '#2F7A4D';    // 완독 - 짙은 초록
const RIBBON_READING = '#C98A2B'; // 읽는 중 - 호박색
const SHOW_PAGE_IN_GROUP = false;
const PAGE_INSET = 5; // 책장: 표지보다 페이지 단면이 좌우로 들어간 깊이(px)
const BOOK_TITLE_FONT = "'Gowun Batang', 'Nanum Myeongjo', serif"; // 책 제목용 한글 세리프 - 붓결이 살아있는 서체
// 책 제목용 한글 명조 폰트 1회 로드
if (typeof document !== 'undefined' && !document.getElementById('font-gowun-batang')) {
  const l = document.createElement('link');
  l.id = 'font-gowun-batang';
  l.rel = 'stylesheet';
  l.href = 'https://fonts.googleapis.com/css2?family=Gowun+Batang:wght@700&display=swap';
  document.head.appendChild(l);
} // 모임 책장에서 남의 진행 페이지 노출 여부 (true면 공개)
const TOWER_BADGE_PALETTE = ['#7C5CC4', '#D97A3D', '#C4544A', '#3E93A0', '#C48A3E'];
// 정확한 우측 90도 측면(옆에서 본 책 두께 단면)용 - 표지 단면에 쓰이는 단색
const COVER_EDGE_COLORS = [
  '#8B5E34', '#3E6B69', '#8A6F45', '#472B1D', '#2E4A66', '#3A5A3A', '#3A3A38', '#663636',
  '#5C3A21', '#264653', '#6B2D3C', '#4B4423', '#3D2B56', '#1F4E4A', '#5A3E5C', '#704214',
  '#B8352E', '#1F5FA8', '#2F8F4E', '#D9A521', // 원색 계열 - 빨강·파랑·초록·노랑
];
// 책마다 미세하게 다른 종이 톤 (같은 크림색이라도 책마다 살짝 다르게)
const PAGE_TONES = [
  { light: '#F3EBD8', dark: '#E7DCC1' },
  { light: '#F0E6D2', dark: '#E2D5B6' },
  { light: '#F5EEE0', dark: '#E9DFC9' },
  { light: '#EEE3CB', dark: '#DFD1AE' },
];
// 페이지 단면에 아주 은은하게 얹는 종이 결 노이즈 텍스처 (외부 라이브러리 없이 SVG data URI로 생성)
const PAGE_NOISE_BG = "url(\"data:image/svg+xml;utantml:parameter name=%3Csvg xmlns='http://www.w3.org/2000/svg' width='60' height='60'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='2' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E\")";
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
const BALLOON_REVEAL_PHOTO_DEFAULT = BALLOON_REVEAL_PHOTO;
// 회원 이름별 생일 사진 매핑 - 이름을 key로 넣어두면 해당 회원 생일에 자동으로 그 사진이 나와요.
// 아직 등록 안 된 회원 생일엔 위 기본 사진(BALLOON_REVEAL_PHOTO_DEFAULT)이 대신 쓰여요.
const BALLOON_REVEAL_PHOTOS = {
  '김시연': 'data:image/jpeg;base64,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',
  '심려현': 'data:image/jpeg;base64,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',
  '이명신': 'data:image/jpeg;base64,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',
  '조혜민': 'data:image/jpeg;base64,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',
  '김경민': 'data:image/jpeg;base64,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',
  '황승모': 'data:image/jpeg;base64,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',
};
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
  // 화면 효과(꽃가루/비/눈) 수동 설정 — 'auto'면 날씨·생일에 따라 자동 결정, 그 외엔 간사가 고른 값을 그대로 씀
  const weatherOverride = settings.find((s) => s.key === 'weatherOverride')?.value || 'auto';
  const setWeatherOverride = async (v) => { await upsertRow('settings', { key: 'weatherOverride', value: v }, 'key'); reload(); };

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
    { key: 'gallery', label: '서재', icon: Library },
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

  // 나주 실시간 날씨 — 생일 꽃가루가 없고 수동 설정도 '자동'일 때만 날씨로 비/눈을 자동 판단
  const [weatherCode, setWeatherCode] = useState(null);
  useEffect(() => {
    let cancelled = false;
    const fetchWeather = () => {
      fetch('https://api.open-meteo.com/v1/forecast?latitude=35.0160&longitude=126.7108&current=weather_code&timezone=Asia%2FSeoul')
        .then((r) => r.json())
        .then((data) => { if (!cancelled) setWeatherCode(data?.current?.weather_code ?? null); })
        .catch(() => {});
    };
    fetchWeather();
    const interval = setInterval(fetchWeather, 30 * 60 * 1000); // 30분마다 갱신
    return () => { cancelled = true; clearInterval(interval); };
  }, []);
  // 화면 효과 우선순위: ① 생일 꽃가루 ② 간사가 수동으로 고른 값 ③ 날씨 자동 판단
  const effectMode = useMemo(() => {
    if (hasBirthdayToday) return 'petal';
    if (weatherOverride && weatherOverride !== 'auto') return weatherOverride === 'off' ? null : weatherOverride;
    if (weatherCode == null) return null;
    const rainCodes = [51, 53, 55, 56, 57, 61, 63, 65, 66, 67, 80, 81, 82, 95, 96, 99];
    const snowCodes = [71, 73, 75, 77, 85, 86];
    if (rainCodes.includes(weatherCode)) return 'rain';
    if (snowCodes.includes(weatherCode)) return 'snow';
    return null;
  }, [hasBirthdayToday, weatherOverride, weatherCode]);
  const weatherPieces = useMemo(() => {
    if (effectMode !== 'rain' && effectMode !== 'snow') return [];
    const count = effectMode === 'rain' ? 55 : 40;
    return Array.from({ length: count }).map((_, i) => ({
      id: i,
      left: Math.random() * 100,
      delay: -(Math.random() * 8).toFixed(2),
      dur: effectMode === 'rain' ? (0.7 + Math.random() * 0.4).toFixed(2) : (7 + Math.random() * 6).toFixed(2),
      size: effectMode === 'rain' ? 1.5 : 3 + Math.random() * 3,
      drift: effectMode === 'rain' ? -30 : Math.round((Math.random() - 0.5) * 60),
    }));
  }, [effectMode]);

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
        @keyframes rainFall {
          0% { transform: translate(0, -10vh); opacity: 0.7; }
          100% { transform: translate(var(--drift), 110vh); opacity: 0.35; }
        }
        @keyframes snowFall {
          0% { transform: translate(0, -10vh) rotate(0deg); opacity: 0.9; }
          100% { transform: translate(var(--drift), 110vh) rotate(180deg); opacity: 0.7; }
        }
      `}</style>
      {effectMode === 'petal' ? (
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
      ) : (effectMode === 'rain' || effectMode === 'snow') && (
        <div className="fixed inset-0 pointer-events-none overflow-hidden" style={{ zIndex: 60 }} aria-hidden="true">
          {weatherPieces.map((p) => (
            effectMode === 'rain' ? (
              <span key={p.id} style={{ position: 'absolute', top: 0, left: `${p.left}%`, width: p.size, height: 16, background: 'linear-gradient(to bottom, rgba(180,210,255,0), rgba(180,210,255,0.55))', '--drift': `${p.drift}px`, animation: `rainFall ${p.dur}s linear ${p.delay}s infinite` }} />
            ) : (
              <span key={p.id} style={{ position: 'absolute', top: 0, left: `${p.left}%`, width: p.size, height: p.size, borderRadius: '50%', background: 'rgba(255,255,255,0.85)', boxShadow: '0 0 4px rgba(255,255,255,0.6)', '--drift': `${p.drift}px`, animation: `snowFall ${p.dur}s linear ${p.delay}s infinite` }} />
            )
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
                  <Lock size={11} />{currentMember ? <><Stamp role={currentMember.role} size={16} tilt={0} />{currentMember.name}</> : '등록된 기기'}
                </span>
              ) : (
                <button onClick={() => (currentMember ? logout() : openLogin())}
                  className="flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold"
                  aria-label={currentMember ? '로그아웃' : '로그인'}
                  style={{ background: currentMember ? BTN_BG : NEUTRAL_BG, color: currentMember ? BTN_TEXT : NEUTRAL_TEXT }}>
                  {currentMember ? <><Stamp role={currentMember.role} size={16} tilt={0} />{currentMember.name}</> : <>로그인</>}
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

        {/* 카테고리 탭 5개를 항상 한 줄에 고정 — 좁은 화면 기준으로 여유 있게 맞추고, 화면이 넓어질수록 지금 크기(패딩14px·폰트14px·아이콘15px)까지 자연스럽게 커짐 */}
        <div className="flex mb-4" style={{ gap: 4 }}>
          {TABS.map((t) => {
            const Icon = t.icon; const active = tab === t.key;
            return (
              <button key={t.key} onClick={() => setTab(t.key)}
                className="flex-1 min-w-0 flex items-center justify-center rounded-full font-semibold"
                style={{
                  background: active ? BTN_BG : CARD_BG, color: active ? BTN_TEXT : MUTE, border: `1px solid ${active ? BTN_BG : LINE}`,
                  gap: 'clamp(2px, 1vw, 6px)', padding: 'clamp(5px, 1.6vw, 8px) clamp(3px, 1.5vw, 14px)', fontSize: 'clamp(10.5px, 2.6vw, 14px)',
                }}>
                <Icon style={{ width: 'clamp(12px, 3.2vw, 15px)', height: 'clamp(12px, 3.2vw, 15px)' }} className="shrink-0" />
                <span className="truncate">{t.label}</span>
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
        {tab === 'admin' && canManageAttendance && <AdminScreen members={sortedMembers} sessions={sessions} checkins={checkins} penaltyRule={penaltyRule} setPenaltyRule={setPenaltyRule} penaltyCompletions={penaltyCompletions} reload={reload} calendarDays={calendarDays} absenceExcuses={absenceExcuses} requestDelete={requestDelete} currentMember={currentMember} weatherOverride={weatherOverride} setWeatherOverride={setWeatherOverride} />}
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
  const birthdayFolksToday = members.filter((m) => m.birthday && mdOf(m.birthday) === todayMd && m.name !== '김태영');

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
        @keyframes photoRevealIn {
          0% { opacity: 0; transform: scale(0.9); }
          100% { opacity: 1; transform: scale(1); }
        }
        @keyframes photoGlowPulse {
          0%, 100% { opacity: 0.75; transform: scale(1); }
          50% { opacity: 1; transform: scale(1.12); }
        }
        @keyframes sparkleTwinkle {
          0%, 100% { opacity: 0.15; transform: scale(0.6) rotate(0deg); }
          50% { opacity: 1; transform: scale(1.15) rotate(15deg); }
        }
      `}</style>
      {birthdayFolksToday.length > 0 && (
        <Card className="text-center">
          <div className="flex items-center justify-center gap-2 mb-1">
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
            <span className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>오늘은 {birthdayFolksToday.map((m) => dispName(m.name, isLoggedIn)).join(', ')}님 생일이에요!</span>
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
          </div>
          <p className="text-sm mb-1" style={{ color: MUTE }}>생일 축하드려요! 행복한 하루 되세요 🎂</p>
          <div className="flex justify-center mt-3">
            <div className="relative" style={{ width: 176, height: 'auto', animation: 'photoRevealIn 1.3s ease-out both' }}>
              {/* 은은하게 숨쉬는 글로우 - 넓게 퍼지는 층 + 밝은 중심층, 두 겹 */}
              <div style={{ position: 'absolute', inset: -34, borderRadius: 50, background: 'radial-gradient(circle, rgba(255,224,120,0.55) 0%, rgba(255,224,120,0) 75%)', animation: 'photoGlowPulse 2.6s ease-in-out infinite', zIndex: -1 }} />
              <div style={{ position: 'absolute', inset: -14, borderRadius: 30, background: 'radial-gradient(circle, rgba(255,244,200,0.85) 0%, rgba(255,213,74,0.4) 45%, rgba(255,213,74,0) 75%)', animation: 'photoGlowPulse 2.6s ease-in-out 0.1s infinite', zIndex: -1 }} />
              <img src={BALLOON_REVEAL_PHOTOS[birthdayFolksToday[0]?.name] || BALLOON_REVEAL_PHOTO_DEFAULT} alt="" style={{ width: 176, height: 'auto', borderRadius: 20, display: 'block', border: '3px solid rgba(255,255,255,0.85)', boxShadow: '0 4px 14px rgba(0,0,0,0.45)' }} />
              {/* 주변을 도는 작은 반짝임 */}
              {[
                { x: -14, y: -8, d: 0 }, { x: 190, y: 4, d: 0.4 }, { x: 200, y: 130, d: 0.9 },
                { x: 176, y: 250, d: 0.2 }, { x: -10, y: 240, d: 0.7 }, { x: -18, y: 120, d: 1.2 },
              ].map((s, i) => (
                <span key={i} style={{ position: 'absolute', left: s.x, top: s.y, fontSize: 16, color: '#FFF4C8', textShadow: '0 0 10px rgba(255,224,120,1), 0 0 18px rgba(255,213,74,0.7)', animation: `sparkleTwinkle 1.8s ease-in-out ${s.d}s infinite` }}>✦</span>
              ))}
            </div>
          </div>
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
              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-1.5 min-w-0">
                  {n.pinned && <Pin size={13} style={{ color: '#EFC94C', flexShrink: 0 }} fill="#EFC94C" />}
                  <h3 className="font-semibold truncate" style={{ color: INK, fontFamily: "'Fraunces', serif", fontSize: 'clamp(13px, 4vw, 16px)' }}>{n.title}</h3>
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
  // 책 검색 (등록 폼)
  const [shareCoverUrl, setShareCoverUrl] = useState('');
  const [bookSearchOpen, setBookSearchOpen] = useState(false);
  const [bookSearchLoading, setBookSearchLoading] = useState(false);
  const [bookSearchResults, setBookSearchResults] = useState([]);
  // 책 검색 (수정 폼)
  const [editShareCoverUrl, setEditShareCoverUrl] = useState('');
  const [editBookSearchOpen, setEditBookSearchOpen] = useState(false);
  const [editBookSearchLoading, setEditBookSearchLoading] = useState(false);
  const [editBookSearchResults, setEditBookSearchResults] = useState([]);
  const [coverUploading, setCoverUploading] = useState(false);
  const [editCoverUploading, setEditCoverUploading] = useState(false);
  const [detailCoverUploading, setDetailCoverUploading] = useState(false);
  // 외부 URL을 그대로 걸지 않고, 우리 저장소로 이미지를 직접 업로드해서 안정적으로 보이게 함
  const uploadCoverImage = async (file, setCoverUrl, setUploading) => {
    if (!file) return;
    if (!file.type.startsWith('image/')) { showToast?.('이미지 파일만 업로드할 수 있어요.', 'error'); return; }
    setUploading(true);
    try {
      const blob = await compressImage(file);
      const path = `book-cover-${uid('bc')}.jpg`;
      const { error: upErr } = await supabase.storage.from('photos').upload(path, blob, { contentType: 'image/jpeg' });
      if (upErr) throw upErr;
      setCoverUrl(publicUrl('photos', path));
    } catch (e) {
      showToast?.('표지 업로드에 실패했어요.', 'error');
    } finally {
      setUploading(false);
    }
  };
  const searchBooks = async (query, setResults, setLoading, setOpen) => {
    if (!query.trim()) return;
    setLoading(true);
    setOpen(true);
    try {
      const res = await fetch(`https://www.googleapis.com/books/v1/volumes?q=${encodeURIComponent(query.trim())}&maxResults=8`);
      const data = await res.json();
      setResults((data.items || []).map((it) => ({
        title: it.volumeInfo?.title || '',
        author: (it.volumeInfo?.authors || []).join(', '),
        publisher: it.volumeInfo?.publisher || '',
        cover: (it.volumeInfo?.imageLinks?.thumbnail || it.volumeInfo?.imageLinks?.smallThumbnail || '').replace(/^http:/, 'https:'),
      })));
    } catch (e) {
      setResults([]);
    }
    setLoading(false);
  };
  const notify = async (memberId, message, linkId) => {
    await insertRow('notifications', { id: uid('nt'), member_id: memberId, message, link_id: linkId, created_at: new Date().toISOString() });
  };
  const submitBookShare = async () => {
    if (!currentMember || !shareTitle.trim()) return;
    await insertRow('book_shares', {
      id: uid('bs'), kind: shareKind, posted_by: currentMember.id,
      book_title: shareTitle.trim(), book_author: shareAuthor.trim() || null, book_publisher: sharePublisher.trim() || null,
      cover_url: shareCoverUrl || null,
      status: 'open', created_at: new Date().toISOString(),
    });
    setShareTitle(''); setShareAuthor(''); setSharePublisher(''); setShareCoverUrl(''); setBookSearchResults([]); setBookSearchOpen(false); setShowShareForm(false);
    await reload();
  };
  const saveShareEdit = async (share) => {
    if (!editShareTitle.trim()) return;
    await updateRow('book_shares', 'id', share.id, { book_title: editShareTitle.trim(), book_author: editShareAuthor.trim() || null, book_publisher: editSharePublisher.trim() || null, cover_url: editShareCoverUrl || null });
    setCoverCache((prev) => { const next = { ...prev }; delete next[share.id]; return next; });
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
    if (s.status === 'requested') return { label: '대여신청중', style: { background: '#332815', color: '#EFC94C' } };
    if (s.status === 'matched') return { label: '대여중', style: { background: '#3A2E10', color: '#EFC94C' } };
    return { label: '반납완료', style: { background: NEUTRAL_BG, color: MUTE, border: `1px solid ${LINE}` } };
  };
  const offers = bookShares.filter((s) => s.kind === 'offer').sort((a, b) => b.created_at.localeCompare(a.created_at));
  const requests = bookShares.filter((s) => s.kind === 'request').sort((a, b) => b.created_at.localeCompare(a.created_at));
  const viewingShare = bookShares.find((s) => s.id === viewingShareId) || null;

  useEffect(() => {
    if (!viewingShare) return;
    if (coverCache[viewingShare.id]) return; // 이미 조회했으면 재요청 안 함
    if (viewingShare.cover_url) { setCoverCache((prev) => ({ ...prev, [viewingShare.id]: viewingShare.cover_url })); return; } // 등록 시 선택해둔 표지가 있으면 그걸 그대로 사용
    let cancelled = false;
    const key = viewingShare.id;
    setCoverCache((prev) => ({ ...prev, [key]: 'loading' }));
    const q = [viewingShare.book_title, viewingShare.book_author].filter(Boolean).map((v) => v.trim()).filter(Boolean).join(' ');
    const query = encodeURIComponent(q);
    fetch(`https://www.googleapis.com/books/v1/volumes?q=${query}&maxResults=1`)
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
  const isSecretary = currentMember?.role === '간사'; // 모임 책장(공개) 등록은 간사만 가능 - 회장·총무는 제외
  const [towerView, setTowerView] = useState('mine'); // 'mine' | 'group'
  const [showTowerAdd, setShowTowerAdd] = useState(false);
  const [towerSettingsOpen, setTowerSettingsOpen] = useState(false); // 켜져 있을 때만 순서 이동·삭제 버튼이 보임 (평소엔 연필만)
  const [towerTitleInput, setTowerTitleInput] = useState('');
  const [towerStartInput, setTowerStartInput] = useState(todayStr());
  const [towerPageInput, setTowerPageInput] = useState('');
  const [towerFinishedInput, setTowerFinishedInput] = useState('');
  const [editingTowerId, setEditingTowerId] = useState(null);
  const [towerPublicInput, setTowerPublicInput] = useState(true);
  const [editingTowerPublic, setEditingTowerPublic] = useState(true);
  const [towerTagInput, setTowerTagInput] = useState('');
  const [editingTowerTag, setEditingTowerTag] = useState('');
  const [towerColorInput, setTowerColorInput] = useState(COVER_EDGE_COLORS[0]); // 내 책장에 등록할 책 표지 색상
  const [editingTowerColor, setEditingTowerColor] = useState(COVER_EDGE_COLORS[0]); // 수정 시 표지 색상
  const addManualTowerEntry = async () => {
    if (!currentMember || !towerTitleInput.trim()) return;
    await insertRow('book_tower_entries', {
      id: uid('bt'), member_id: currentMember.id, book_title: towerTitleInput.trim(),
      start_date: towerStartInput || null, current_page: towerPageInput ? parseInt(towerPageInput, 10) : null, finished_date: towerFinishedInput || null,
      color: towerColorInput, source_share_id: null, created_at: new Date().toISOString(), sort_order: Date.now(),
      is_public: isSecretary ? towerPublicInput : false, // 모임 책장 공개는 간사만 가능 - 그 외 회원 글은 항상 내 책장에만
      event_tag: towerTagInput.trim() || null,
      owner_name_override: null,
    });
    setTowerTitleInput(''); setTowerStartInput(todayStr()); setTowerPageInput(''); setTowerFinishedInput(''); setTowerPublicInput(true); setTowerTagInput(''); setTowerColorInput(COVER_EDGE_COLORS[Math.floor(Math.random() * COVER_EDGE_COLORS.length)]); setShowTowerAdd(false);
    await reload();
  };
  const saveTowerEdit = async (entry) => {
    try {
      const { data, error } = await supabase.from('book_tower_entries')
        .update({
          start_date: towerStartInput || null, current_page: towerPageInput ? parseInt(towerPageInput, 10) : null, finished_date: towerFinishedInput || null,
          is_public: isSecretary ? editingTowerPublic : entry.is_public,
          event_tag: editingTowerTag.trim() || null,
          owner_name_override: entry.owner_name_override,
          color: editingTowerColor,
        })
        .eq('id', entry.id)
        .select();
      if (error) throw error;
      if (!data || data.length === 0) throw new Error('저장 권한이 없거나 항목을 찾지 못했어요 (관리자에게 문의해주세요).');
      setEditingTowerId(null);
      await reload();
    } catch (e) {
      showToast?.('저장에 실패했어요: ' + (e?.message || '알 수 없는 오류'), 'error');
    }
  };
  const startTowerEdit = (entry) => {
    setEditingTowerId(entry.id);
    setTowerStartInput(entry.start_date || '');
    setTowerPageInput(entry.current_page ? String(entry.current_page) : '');
    setTowerFinishedInput(entry.finished_date || '');
    setEditingTowerPublic(entry.is_public !== false);
    setEditingTowerTag(entry.event_tag || '');
    setEditingTowerColor(entry.color || COVER_EDGE_COLORS[0]);
  };
  const removeTowerEntry = async (entryId) => { await deleteRow('book_tower_entries', 'id', entryId); await reload(); };
  const towerSortKey = (t) => t.finished_date || t.start_date || t.created_at.slice(0, 10);
  // 순서를 손으로 바꾼 적이 있으면 sort_order를, 없으면 날짜를 기준으로 삼음 (둘 다 밀리초 단위 숫자라 섞여도 자연스럽게 정렬됨)
  const towerEffectiveOrder = (t) => t.sort_order != null ? t.sort_order : (new Date(towerSortKey(t)).getTime() || 0);
  const towerCompare = (a, b) => towerEffectiveOrder(a) - towerEffectiveOrder(b);
  const myTower = currentMember ? bookTowerEntries.filter((t) => t.member_id === currentMember.id).sort(towerCompare) : [];
  // 비공개(is_public === false)로 설정한 책은 모임 책장에서는 빠지고 본인 책장에서만 보임. 기존 데이터(is_public 필드 없음)는 공개로 취급
  const groupTower = bookTowerEntries.filter((t) => t.is_public !== false).sort(towerCompare);
  // 목록 안에서 책탑 순서를 위/아래로 한 칸씩 바꿈 (쌓인 순서상 뒤 항목=화면 위쪽, 앞 항목=화면 아래쪽)
  const moveTowerItem = async (list, entry, direction) => {
    const idx = list.findIndex((x) => x.id === entry.id);
    const swapIdx = direction === 'up' ? idx + 1 : idx - 1;
    if (idx === -1 || swapIdx < 0 || swapIdx >= list.length) return;
    const other = list[swapIdx];
    const a = towerEffectiveOrder(entry);
    const b = towerEffectiveOrder(other);
    try {
      await updateRow('book_tower_entries', 'id', entry.id, { sort_order: b });
      await updateRow('book_tower_entries', 'id', other.id, { sort_order: a });
      await reload();
    } catch (err) {
      // sort_order 컬럼이 아직 Supabase의 book_tower_entries 테이블에 없으면 여기서 실패함
      showToast('순서 변경에 실패했어요 — book_tower_entries 테이블에 sort_order 컬럼이 있는지 확인해주세요.', 'error');
    }
  };

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
          <div className="rounded-2xl p-3 mb-4 space-y-2" style={{ background: FORM_PANEL_BG, border: `1.5px solid ${FORM_PANEL_BORDER}`, boxShadow: 'inset 0 1px 3px rgba(0,0,0,0.3)' }}>
            <div className="flex items-center gap-1.5 text-xs font-bold mb-0.5" style={{ color: MUTE }}><Pencil size={11} /> 새 글 작성</div>
            <div className="flex gap-2">
              <button onClick={() => setShareKind('offer')} className="flex-1 rounded-xl py-2 text-xs font-semibold" style={{ background: shareKind === 'offer' ? BTN_BG : NEUTRAL_BG, color: shareKind === 'offer' ? BTN_TEXT : NEUTRAL_TEXT }}>빌려줄까요?</button>
              <button onClick={() => setShareKind('request')} className="flex-1 rounded-xl py-2 text-xs font-semibold" style={{ background: shareKind === 'request' ? BTN_BG : NEUTRAL_BG, color: shareKind === 'request' ? BTN_TEXT : NEUTRAL_TEXT }}>빌려주실 수 있나요?</button>
            </div>
            <div className="relative">
              <div className="flex gap-2">
                <input value={shareTitle} onChange={(e) => { setShareTitle(e.target.value); setShareCoverUrl(''); }} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); searchBooks(shareTitle, setBookSearchResults, setBookSearchLoading, setBookSearchOpen); } }} placeholder="책 제목 (필수)" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                <button onClick={() => searchBooks(shareTitle, setBookSearchResults, setBookSearchLoading, setBookSearchOpen)} disabled={!shareTitle.trim()} className="shrink-0 rounded-xl px-3 disabled:opacity-40" style={{ background: NEUTRAL_BG }} aria-label="책 검색"><Search size={16} style={{ color: NEUTRAL_TEXT }} /></button>
              </div>
              {bookSearchOpen && (
                <div className="absolute z-10 left-0 right-0 mt-1 rounded-xl border overflow-hidden" style={{ background: CARD_BG, borderColor: LINE, maxHeight: 260, overflowY: 'auto' }}>
                  <div className="flex items-center justify-between px-2 py-1" style={{ borderBottom: `1px solid ${ROW_LINE}` }}>
                    <span className="text-[11px]" style={{ color: MUTE }}>{bookSearchLoading ? '검색 중…' : `검색 결과 ${bookSearchResults.length}건`}</span>
                    <button onClick={() => setBookSearchOpen(false)} className="p-1" aria-label="검색 결과 닫기"><X size={12} style={{ color: MUTE }} /></button>
                  </div>
                  {!bookSearchLoading && bookSearchResults.length === 0 && (
                    <div className="px-3 py-3 text-xs space-y-2" style={{ color: MUTE }}>
                      <div>검색 결과가 없어요.</div>
                      <a href={`https://www.google.com/search?tbm=isch&q=${encodeURIComponent((shareTitle || '') + ' 책 표지')}`} target="_blank" rel="noreferrer" className="inline-block underline" style={{ color: NEUTRAL_TEXT }}>구글 이미지에서 표지 찾아보기 →</a>
                    </div>
                  )}
                  {bookSearchResults.map((r, i) => (
                    <button key={i} onClick={() => { setShareTitle(r.title); setShareAuthor(r.author); setSharePublisher(r.publisher); setShareCoverUrl(r.cover); setBookSearchOpen(false); }} className="w-full flex items-center gap-2.5 px-2.5 py-2 text-left" style={{ borderBottom: i < bookSearchResults.length - 1 ? `1px solid ${ROW_LINE}` : 'none' }}>
                      {r.cover ? <img src={r.cover} alt="" className="rounded shrink-0" style={{ width: 32, height: 46, objectFit: 'cover' }} /> : <div className="rounded shrink-0" style={{ width: 32, height: 46, background: NEUTRAL_BG }} />}
                      <div className="min-w-0">
                        <div className="text-xs font-semibold truncate" style={{ color: INK }}>{r.title}</div>
                        <div className="text-[11px] truncate" style={{ color: MUTE }}>{[r.author, r.publisher].filter(Boolean).join(' · ')}</div>
                      </div>
                    </button>
                  ))}
                </div>
              )}
            </div>
            {shareCoverUrl && (
              <div className="flex items-center gap-2">
                <img src={shareCoverUrl} alt="" className="rounded shadow-sm" style={{ width: 32, height: 46, objectFit: 'cover' }} />
                <span className="text-[11px]" style={{ color: MUTE }}>표지가 선택됐어요</span>
                <button onClick={() => setShareCoverUrl('')} className="text-[11px] underline" style={{ color: MUTE }}>선택 해제</button>
              </div>
            )}
            <label className={`flex items-center justify-center gap-1.5 text-xs font-semibold rounded-xl border-2 border-dashed py-2.5 cursor-pointer ${coverUploading ? 'opacity-60' : ''}`} style={{ borderColor: LINE, color: NEUTRAL_TEXT }}>
              <ImageIcon size={14} />
              {coverUploading ? '업로드 중…' : '표지 사진 올리기 (구글 이미지 캡처 등)'}
              <input type="file" accept="image/*" className="hidden" disabled={coverUploading} onChange={(e) => { const f = e.target.files[0]; e.target.value = ''; uploadCoverImage(f, setShareCoverUrl, setCoverUploading); }} />
            </label>
            <div className="grid grid-cols-2 gap-2">
              <input value={shareAuthor} onChange={(e) => setShareAuthor(e.target.value)} placeholder="저자 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
              <input value={sharePublisher} onChange={(e) => setSharePublisher(e.target.value)} placeholder="출판사 (선택)" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            </div>
            <PrimaryBtn onClick={submitBookShare} icon={Plus}>등록</PrimaryBtn>
          </div>
        )}
        {!currentMember && !showShareForm && <p className="text-xs mb-2" style={{ color: MUTE }}>상단에서 본인을 먼저 선택해야 글을 올릴 수 있어요.</p>}
        {showShareForm && <div className="text-[10px] font-semibold mb-2" style={{ color: MUTE }}>등록된 글 목록</div>}
        <div className="inline-flex items-center gap-1.5 mb-1.5 text-xs font-bold" style={{ color: SHARE_OFFER_COLOR }}><span style={{ fontSize: 13, lineHeight: 1 }} role="img" aria-label="손 내미는 사람">💁‍♀️</span> 빌려줄까요? ({offers.length})</div>
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
        <div className="inline-flex items-center gap-1.5 mb-1.5 text-xs font-bold" style={{ color: SHARE_REQUEST_COLOR }}>🙋 빌려주실 수 있나요? ({requests.length})</div>
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
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 overflow-y-auto" style={{ background: 'rgba(0,0,0,0.7)' }} onClick={() => setViewingShareId(null)}>
            <div className="w-full max-w-sm rounded-2xl border p-5 my-auto" style={{ background: CARD_BG, borderColor: LINE }} onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between mb-3">
                <span className="text-[11px] rounded-full px-2 py-0.5 font-semibold" style={{ background: viewingShare.kind === 'offer' ? SHARE_OFFER_BG : SHARE_REQUEST_BG, color: viewingShare.kind === 'offer' ? SHARE_OFFER_COLOR : SHARE_REQUEST_COLOR }}>{viewingShare.kind === 'offer' ? '빌려줄까요?' : '빌려주실 수 있나요?'}</span>
                <div className="flex items-center gap-2">
                  {canEditPost && !editingShare && (
                    <button onClick={() => { setEditingShare(true); setEditShareTitle(viewingShare.book_title); setEditShareAuthor(viewingShare.book_author || ''); setEditSharePublisher(viewingShare.book_publisher || ''); setEditShareCoverUrl(viewingShare.cover_url || ''); }} aria-label="글 수정"><Pencil size={14} style={{ color: MUTE }} /></button>
                  )}
                  {canEditPost && !editingShare && (
                    <button onClick={() => requestDelete(() => deleteBookShare(viewingShare), '이 게시글을 삭제할까요? 대여 기록도 함께 사라져요.')} aria-label="글 삭제"><Trash2 size={14} style={{ color: '#F0A87C' }} /></button>
                  )}
                  <button onClick={() => setViewingShareId(null)} aria-label="닫기"><X size={16} style={{ color: MUTE }} /></button>
                </div>
              </div>
              {editingShare ? (
                <div className="space-y-2 mb-3">
                  <div className="relative">
                    <div className="flex gap-2">
                      <input value={editShareTitle} onChange={(e) => { setEditShareTitle(e.target.value); }} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); searchBooks(editShareTitle, setEditBookSearchResults, setEditBookSearchLoading, setEditBookSearchOpen); } }} placeholder="책 제목" className="flex-1 rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                      <button onClick={() => searchBooks(editShareTitle, setEditBookSearchResults, setEditBookSearchLoading, setEditBookSearchOpen)} disabled={!editShareTitle.trim()} className="shrink-0 rounded-xl px-3 disabled:opacity-40" style={{ background: NEUTRAL_BG }} aria-label="책 검색"><Search size={16} style={{ color: NEUTRAL_TEXT }} /></button>
                    </div>
                    {editBookSearchOpen && (
                      <div className="absolute z-10 left-0 right-0 mt-1 rounded-xl border overflow-hidden" style={{ background: CARD_BG, borderColor: LINE, maxHeight: 260, overflowY: 'auto' }}>
                        <div className="flex items-center justify-between px-2 py-1" style={{ borderBottom: `1px solid ${ROW_LINE}` }}>
                          <span className="text-[11px]" style={{ color: MUTE }}>{editBookSearchLoading ? '검색 중…' : `검색 결과 ${editBookSearchResults.length}건`}</span>
                          <button onClick={() => setEditBookSearchOpen(false)} className="p-1" aria-label="검색 결과 닫기"><X size={12} style={{ color: MUTE }} /></button>
                        </div>
                        {!editBookSearchLoading && editBookSearchResults.length === 0 && (
                          <div className="px-3 py-3 text-xs space-y-2" style={{ color: MUTE }}>
                            <div>검색 결과가 없어요.</div>
                            <a href={`https://www.google.com/search?tbm=isch&q=${encodeURIComponent((editShareTitle || '') + ' 책 표지')}`} target="_blank" rel="noreferrer" className="inline-block underline" style={{ color: NEUTRAL_TEXT }}>구글 이미지에서 표지 찾아보기 →</a>
                          </div>
                        )}
                        {editBookSearchResults.map((r, i) => (
                          <button key={i} onClick={() => { setEditShareTitle(r.title); setEditShareAuthor(r.author); setEditSharePublisher(r.publisher); setEditShareCoverUrl(r.cover); setEditBookSearchOpen(false); }} className="w-full flex items-center gap-2.5 px-2.5 py-2 text-left" style={{ borderBottom: i < editBookSearchResults.length - 1 ? `1px solid ${ROW_LINE}` : 'none' }}>
                            {r.cover ? <img src={r.cover} alt="" className="rounded shrink-0" style={{ width: 32, height: 46, objectFit: 'cover' }} /> : <div className="rounded shrink-0" style={{ width: 32, height: 46, background: NEUTRAL_BG }} />}
                            <div className="min-w-0">
                              <div className="text-xs font-semibold truncate" style={{ color: INK }}>{r.title}</div>
                              <div className="text-[11px] truncate" style={{ color: MUTE }}>{[r.author, r.publisher].filter(Boolean).join(' · ')}</div>
                            </div>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                  {editShareCoverUrl && (
                    <div className="flex items-center gap-2">
                      <img src={editShareCoverUrl} alt="" className="rounded shadow-sm" style={{ width: 32, height: 46, objectFit: 'cover' }} />
                      <span className="text-[11px]" style={{ color: MUTE }}>표지가 선택됐어요</span>
                      <button onClick={() => setEditShareCoverUrl('')} className="text-[11px] underline" style={{ color: MUTE }}>선택 해제</button>
                    </div>
                  )}
                  <label className={`flex items-center justify-center gap-1.5 text-xs font-semibold rounded-xl border-2 border-dashed py-2.5 cursor-pointer ${editCoverUploading ? 'opacity-60' : ''}`} style={{ borderColor: LINE, color: NEUTRAL_TEXT }}>
                    <ImageIcon size={14} />
                    {editCoverUploading ? '업로드 중…' : '표지 사진 올리기 (구글 이미지 캡처 등)'}
                    <input type="file" accept="image/*" className="hidden" disabled={editCoverUploading} onChange={(e) => { const f = e.target.files[0]; e.target.value = ''; uploadCoverImage(f, setEditShareCoverUrl, setEditCoverUploading); }} />
                  </label>
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
                      <img src={coverCache[viewingShare.id]} alt="" className="rounded-lg shadow-md" style={{ height: 408, width: 'auto', maxWidth: '100%' }} />
                    </div>
                  )}
                  {coverCache[viewingShare.id] === 'none' && canEditPost && (
                    <div className="flex justify-center mb-3">
                      <label className={`flex items-center justify-center gap-1.5 text-xs font-semibold rounded-xl border-2 border-dashed px-4 py-2.5 cursor-pointer ${detailCoverUploading ? 'opacity-60' : ''}`} style={{ borderColor: LINE, color: NEUTRAL_TEXT }}>
                        <ImageIcon size={14} />
                        {detailCoverUploading ? '업로드 중…' : '표지 사진 올리기'}
                        <input type="file" accept="image/*" className="hidden" disabled={detailCoverUploading} onChange={(e) => {
                          const f = e.target.files[0]; e.target.value = '';
                          uploadCoverImage(f, async (url) => {
                            await updateRow('book_shares', 'id', viewingShare.id, { cover_url: url });
                            setCoverCache((prev) => ({ ...prev, [viewingShare.id]: url }));
                            await reload();
                          }, setDetailCoverUploading);
                        }} />
                      </label>
                    </div>
                  )}
                  {coverCache[viewingShare.id] === 'loading' && (
                    <div className="flex justify-center mb-3">
                      <div className="rounded-lg animate-pulse" style={{ height: 408, width: 278, background: NEUTRAL_BG }} />
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
                    <div className="flex justify-end">
                      <PrimaryBtn onClick={() => { requestShare(viewingShare); setViewingShareId(null); }} icon={Check}>대여신청</PrimaryBtn>
                    </div>
                  ) : currentMember?.id === viewingShare.posted_by ? (
                    <p className="text-xs" style={{ color: MUTE }}>다른 회원의 응답을 기다리는 중이에요.</p>
                  ) : null}
                </>
              )}
              {viewingShare.status === 'requested' && !editingShare && (
                <div className="space-y-2 pt-2" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <span className="inline-flex items-center text-[11px] rounded-full px-2 py-0.5 font-semibold" style={shareStatusInfo(viewingShare).style}>{shareStatusInfo(viewingShare).label}</span>
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
            </div>
          </div>
        );
      })()}

      {/* ---------- 책탑 ---------- */}
      <Card>
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-1.5 text-sm font-semibold" style={{ color: INK }}>📚 북적북적</div>
          <div className="flex items-center gap-1.5">
            {currentMember && <button onClick={() => setShowTowerAdd((v) => !v)} className="text-xs rounded-full px-3 py-1.5 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>{showTowerAdd ? '취소' : '+ 책 추가'}</button>}
            <button onClick={() => setTowerSettingsOpen((v) => !v)} className="p-2 rounded-full" style={{ background: towerSettingsOpen ? BTN_BG : NEUTRAL_BG, color: towerSettingsOpen ? BTN_TEXT : NEUTRAL_TEXT }} aria-label="책탑 설정 (순서 변경·삭제)"><Settings2 size={14} /></button>
          </div>
        </div>
        {showTowerAdd && (
          <div className="rounded-2xl p-3 mb-3 space-y-2" style={{ background: FORM_PANEL_BG, border: `1.5px solid ${FORM_PANEL_BORDER}`, boxShadow: 'inset 0 1px 3px rgba(0,0,0,0.3)' }}>
            <div className="flex items-center justify-between mb-0.5">
              <div className="flex items-center gap-1.5 text-xs font-bold" style={{ color: MUTE }}><Pencil size={11} /> 새 책 추가</div>
              {isSecretary && (
                <label className="flex items-center gap-1.5 text-[11px]" style={{ color: MUTE }}>
                  <input type="checkbox" checked={towerPublicInput} onChange={(e) => setTowerPublicInput(e.target.checked)} />
                  모임 책장에도 공개
                </label>
              )}
            </div>
            <input value={towerTitleInput} onChange={(e) => setTowerTitleInput(e.target.value)} placeholder="책 제목" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <div className="grid grid-cols-2 gap-2">
              <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>읽기 시작일</div><input type="date" value={towerStartInput} onChange={(e) => setTowerStartInput(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
              <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>다 읽은 날 (선택)</div><input type="date" value={towerFinishedInput} onChange={(e) => setTowerFinishedInput(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            </div>
            <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>현재 읽고 있는 페이지 (선택)</div><input type="number" value={towerPageInput} onChange={(e) => setTowerPageInput(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>행사/토론회 태그 (선택)</div><input value={towerTagInput} onChange={(e) => setTowerTagInput(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            <div>
              <div className="text-[10px] mb-1" style={{ color: MUTE }}>표지 색상</div>
              <div className="flex flex-wrap gap-1.5">
                {COVER_EDGE_COLORS.map((c) => (
                  <button key={c} type="button" onClick={() => setTowerColorInput(c)} aria-label={`표지 색상 ${c}`}
                    className="rounded-full shrink-0" style={{ width: 24, height: 24, background: c, border: towerColorInput === c ? `2px solid ${BTN_BG}` : `1px solid ${LINE}`, boxShadow: towerColorInput === c ? '0 0 0 2px rgba(242,238,227,0.15)' : 'none' }} />
                ))}
              </div>
            </div>
            {!isSecretary && (
              <p className="text-[11px]" style={{ color: MUTE }}>* 모임 책장 공개는 간사만 가능해요. 이 책은 내 책장에만 기록돼요.</p>
            )}
            <PrimaryBtn onClick={addManualTowerEntry} icon={Plus}>추가</PrimaryBtn>
          </div>
        )}
        {showTowerAdd && <div className="text-[10px] font-semibold mb-2" style={{ color: MUTE }}>쌓인 책 목록</div>}
        <div className="flex gap-2 mb-3">
          <button onClick={() => setTowerView('mine')} className="flex-1 rounded-xl py-1.5 text-xs font-semibold" style={{ background: towerView === 'mine' ? BTN_BG : NEUTRAL_BG, color: towerView === 'mine' ? BTN_TEXT : NEUTRAL_TEXT }}>내 책장 ({myTower.length})</button>
          <button onClick={() => setTowerView('group')} className="flex-1 rounded-xl py-1.5 text-xs font-semibold" style={{ background: towerView === 'group' ? BTN_BG : NEUTRAL_BG, color: towerView === 'group' ? BTN_TEXT : NEUTRAL_TEXT }}>모임 책장 ({groupTower.length})</button>
        </div>
        {(() => {
          const list = towerView === 'mine' ? myTower : groupTower;
          if (list.length === 0) return <p className="text-xs text-center py-6" style={{ color: MUTE }}>아직 쌓인 책이 없어요.</p>;
          return (
            <div className="flex flex-col-reverse gap-1.5 max-h-[28rem] overflow-y-auto pr-1">
              {list.map((t, idx) => {
                const owner = towerView === 'group' ? members.find((m) => m.id === t.member_id) : null;
                const isMine = towerView === 'mine' && currentMember;
                const canReorder = towerView === 'mine' ? !!isMine : isSecretary; // 내 책장은 본인이, 모임 책장은 간사만 순서 변경 가능
                const canMoveUp = canReorder && idx < list.length - 1; // 배열 뒤쪽일수록 화면 위쪽에 쌓이므로 '위로'는 다음 인덱스와 교체
                const canMoveDown = canReorder && idx > 0;
                const isEditing = editingTowerId === t.id;
                const hash = t.id.split('').reduce((s, c) => s + c.charCodeAt(0), 0);
                const edgeColor = t.color || COVER_EDGE_COLORS[hash % COVER_EDGE_COLORS.length]; // 등록 시 고른 색상 우선, 없으면(기존 항목) 해시 기반 자동 색상
                const isDone = !!t.finished_date;
                const ribbonStatusColor = isDone ? RIBBON_DONE : RIBBON_READING;
                const ribbonLabel = isDone ? '완독' : '읽는 중';
                // 읽는 중일 때만 페이지를 작은 동그라미로 표시 (모임 책장은 SHOW_PAGE_IN_GROUP 설정 따름)
                const showPage = !isDone && t.current_page && (towerView === 'mine' || SHOW_PAGE_IN_GROUP);
                if (isEditing) {
                  return (
                    <div key={t.id} className="rounded-2xl p-3 space-y-2" style={{ background: FORM_PANEL_BG, border: `1.5px solid ${FORM_PANEL_BORDER}`, boxShadow: 'inset 0 1px 3px rgba(0,0,0,0.3)' }}>
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-1.5 text-xs font-bold" style={{ color: MUTE }}><Pencil size={11} /> 책 정보 수정</div>
                        {isSecretary && (
                          <label className="flex items-center gap-1.5 text-[11px]" style={{ color: MUTE }}>
                            <input type="checkbox" checked={editingTowerPublic} onChange={(e) => setEditingTowerPublic(e.target.checked)} />
                            모임 책장에 공개
                          </label>
                        )}
                      </div>
                      <div className="text-sm font-bold truncate" style={{ color: INK }}>{t.book_title}</div>
                      <div className="grid grid-cols-2 gap-2">
                        <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>읽기 시작일</div><input type="date" value={towerStartInput} onChange={(e) => setTowerStartInput(e.target.value)} className="w-full rounded-xl border px-2.5 py-1.5 text-[13px] outline-none" style={inputStyle} aria-label="읽기 시작일" /></div>
                        <div>
                          <div className="text-[10px] mb-1" style={{ color: MUTE }}>다 읽은 날</div>
                          <div className="flex gap-1 items-center">
                            <input type="date" value={towerFinishedInput} onChange={(e) => setTowerFinishedInput(e.target.value)} className="w-full rounded-xl border px-2.5 py-1.5 text-[13px] outline-none flex-1 min-w-0" style={inputStyle} aria-label="다 읽은 날" />
                            {towerFinishedInput && (
                              <button onClick={() => setTowerFinishedInput('')} className="shrink-0 p-1.5" aria-label="다 읽은 날 지우기"><X size={13} style={{ color: MUTE }} /></button>
                            )}
                          </div>
                        </div>
                      </div>
                      <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>현재 읽고 있는 페이지</div><input type="number" value={towerPageInput} onChange={(e) => setTowerPageInput(e.target.value)} className="w-full rounded-xl border px-2.5 py-1.5 text-[13px] outline-none" style={inputStyle} /></div>
                      <div><div className="text-[10px] mb-1" style={{ color: MUTE }}>행사/토론회 태그</div><input value={editingTowerTag} onChange={(e) => setEditingTowerTag(e.target.value)} className="w-full rounded-xl border px-2.5 py-1.5 text-[13px] outline-none" style={inputStyle} /></div>
                      <div>
                        <div className="text-[10px] mb-1" style={{ color: MUTE }}>표지 색상</div>
                        <div className="flex flex-wrap gap-1.5">
                          {COVER_EDGE_COLORS.map((c) => (
                            <button key={c} type="button" onClick={() => setEditingTowerColor(c)} aria-label={`표지 색상 ${c}`}
                              className="rounded-full shrink-0" style={{ width: 22, height: 22, background: c, border: editingTowerColor === c ? `2px solid ${BTN_BG}` : `1px solid ${LINE}`, boxShadow: editingTowerColor === c ? '0 0 0 2px rgba(242,238,227,0.15)' : 'none' }} />
                          ))}
                        </div>
                      </div>
                      <div className="flex gap-1.5 items-center pt-1">
                        <button onClick={() => saveTowerEdit(t)} className="flex-1 text-xs rounded-full py-2 font-semibold" style={{ background: '#1E1C16', color: '#F2EEE3' }}>저장</button>
                        <button onClick={() => setEditingTowerId(null)} className="flex-1 text-xs rounded-full py-2 font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>취소</button>
                      </div>
                    </div>
                  );
                }
                const jitterX = ((hash % 7) - 3) * 5; // 중심에서 좌우로 살짝씩만 어긋나게 (쌓인 더미의 중심은 유지)
                const lengthInset = 6 + (hash % 3) * 4; // 책마다 길이(폭)도 살짝 다르게 - 항상 가운데 기준으로 좁아짐 (제목이 잘리지 않도록 여백을 좁게)
                const microY = (hash % 3) - 1; // -1~1px, 실제로 쌓았을 때 생기는 미세한 높이 오차
                const barH = [42, 47, 53, 58][hash % 4] + (t.event_tag ? 8 : 0); // 책마다 두께를 다르게, 태그 있으면 한 줄만큼만 여유
                const edgeH = Math.max(3, Math.round(barH * 0.112)); // 위아래 표지 두께 (기존의 70%)
                const tone = PAGE_TONES[hash % PAGE_TONES.length];
                return (
                  <div key={t.id} style={{ height: barH, marginLeft: Math.max(2, lengthInset + jitterX), marginRight: Math.max(2, lengthInset - jitterX), transform: `translateY(${microY}px)`, borderRadius: 4, overflow: 'hidden', filter: 'drop-shadow(0 3px 4px rgba(0,0,0,0.45)) drop-shadow(0 1px 1px rgba(0,0,0,0.4))' }}>
                    <div className="relative w-full h-full">
                      {/* 위/아래 - 실제 양장본 표지: 재질감 있는 그라데이션 + 페이지와 맞닿는 경계 그림자 */}
                      <div className="absolute inset-x-0 top-0 pointer-events-none" style={{ height: edgeH, borderRadius: '4px 4px 2px 2px', background: `linear-gradient(180deg, rgba(255,255,255,0.30) 0%, ${edgeColor} 55%, rgba(0,0,0,0.15) 100%)`, boxShadow: 'inset 0 -2px 3px -1px rgba(0,0,0,0.35)' }} />
                      <div className="absolute inset-x-0 bottom-0 pointer-events-none" style={{ height: edgeH, borderRadius: '2px 2px 4px 4px', background: `linear-gradient(0deg, rgba(0,0,0,0.35) 0%, ${edgeColor} 55%, rgba(255,255,255,0.10) 100%)`, boxShadow: 'inset 0 2px 3px -1px rgba(0,0,0,0.35)' }} />
                      {/* 가운데 - 종이 페이지 단면: 촘촘한 결 + 은은한 볼륨감 + 종이 노이즈, 표지가 좌우로 1px 살짝 돌출 */}
                      <div className="absolute pointer-events-none" style={{ top: edgeH, bottom: edgeH, left: PAGE_INSET, right: PAGE_INSET, borderRadius: 3,
                        background: `linear-gradient(180deg, rgba(0,0,0,0.16) 0%, rgba(0,0,0,0) 18%, rgba(0,0,0,0) 82%, rgba(0,0,0,0.20) 100%),
                          repeating-linear-gradient(180deg, ${tone.light} 0px, ${tone.light} 1.4px, ${tone.dark} 1.4px, ${tone.dark} 2.1px)` }} />
                      <div className="absolute pointer-events-none" style={{ top: edgeH, bottom: edgeH, left: PAGE_INSET, right: PAGE_INSET, borderRadius: 3, backgroundImage: PAGE_NOISE_BG, backgroundSize: '60px 60px', opacity: 0.05, mixBlendMode: 'multiply' }} />
                      {/* 좌우 끝 - 페이지 묶음이 표지 안쪽으로 둥글게 말려 들어가는 음영 (좌우 대칭) */}
                      <div className="absolute pointer-events-none" style={{ top: edgeH, bottom: edgeH, left: PAGE_INSET, right: PAGE_INSET, borderRadius: 3,
                        background: 'linear-gradient(90deg, rgba(0,0,0,0.28) 0px, rgba(0,0,0,0.08) 4px, rgba(0,0,0,0) 9px, rgba(0,0,0,0) calc(100% - 9px), rgba(0,0,0,0.08) calc(100% - 4px), rgba(0,0,0,0.28) 100%)',
                        boxShadow: 'inset 1px 0 0 rgba(255,255,255,0.25), inset -1px 0 0 rgba(255,255,255,0.25)' }} />
                      {/* 오른쪽 끝 - 상태 리본 (완독=초록 / 읽는중=호박색) */}
                      <div className="absolute pointer-events-none" style={{ top: -1, right: 12, width: 9, height: Math.round(barH * 0.62), background: `linear-gradient(90deg, ${ribbonStatusColor} 0%, ${ribbonStatusColor}cc 100%)`, clipPath: 'polygon(0 0, 100% 0, 100% 100%, 50% calc(100% - 4px), 0 100%)', boxShadow: '1px 1px 2px rgba(0,0,0,0.3)' }} />
                      <div className="relative h-full flex items-center justify-between gap-1.5 pl-3" style={{ paddingRight: 20 }}>
                        {/* 왼쪽 - 행사 태그(캡션) 위, 제목·등록자 아래로 한 덩어리 */}
                        <div className="min-w-0 flex flex-col justify-center" style={{ gap: 2 }}>
                          {t.event_tag && (
                            <span className="truncate" style={{ fontSize: 9, lineHeight: '11px', fontWeight: 600, color: '#9C7B4A', letterSpacing: '0.15px' }}>{t.event_tag}</span>
                          )}
                          <div className="min-w-0 flex items-baseline gap-1">
                            {t.is_public === false && <Lock size={10} style={{ color: '#6B5B3E', alignSelf: 'center' }} />}
                            <span className="truncate min-w-0" style={{ fontFamily: BOOK_TITLE_FONT, fontSize: 13.5, fontWeight: 700, lineHeight: '17px', color: '#2A2015', letterSpacing: '0px', textShadow: '0 1px 0 rgba(255,255,255,0.3)' }}>{t.book_title}</span>
                            {(t.owner_name_override != null ? t.owner_name_override : (owner ? dispName(owner.name, isLoggedIn) : '')) && (
                              <span className="text-[10px] truncate shrink-0" style={{ color: '#8A7355', fontWeight: 500 }}>
                                {t.owner_name_override != null ? t.owner_name_override : (owner ? dispName(owner.name, isLoggedIn) : '')}
                              </span>
                            )}
                          </div>
                        </div>
                        <div className="flex items-center gap-1.5 shrink-0">
                          {/* 상태 - 읽는 중이면 쪽수를 바로 아래 줄에, 완독이면 한 줄만 */}
                          <div className="flex flex-col items-end" style={{ gap: 1 }}>
                            <span style={{ fontSize: 10, lineHeight: '11px', fontWeight: 700, color: ribbonStatusColor, letterSpacing: '0.1px' }}>{ribbonLabel}</span>
                            {showPage && (
                              <span style={{ fontSize: 9, lineHeight: '10px', color: '#8A6A3F', fontWeight: 600, fontFamily: "'IBM Plex Mono', monospace", fontVariantNumeric: 'tabular-nums' }} aria-label={`현재 ${t.current_page}페이지`}>{t.current_page}쪽</span>
                            )}
                          </div>
                          {canReorder && towerSettingsOpen && (
                            <div className="flex flex-col rounded-lg overflow-hidden" style={{ gap: 2, background: NEUTRAL_BG }}>
                              <button onClick={() => moveTowerItem(list, t, 'up')} disabled={!canMoveUp} className="flex items-center justify-center" style={{ width: 28, height: 18, opacity: canMoveUp ? 1 : 0.25 }} aria-label="위로 이동"><ChevronUp size={14} style={{ color: '#6B5B3E' }} /></button>
                              <button onClick={() => moveTowerItem(list, t, 'down')} disabled={!canMoveDown} className="flex items-center justify-center" style={{ width: 28, height: 18, opacity: canMoveDown ? 1 : 0.25 }} aria-label="아래로 이동"><ChevronDown size={14} style={{ color: '#6B5B3E' }} /></button>
                            </div>
                          )}
                          {isMine && (
                            <div className="flex items-center gap-1">
                              <button onClick={() => startTowerEdit(t)} className="p-0.5" aria-label="책탑 항목 수정"><Pencil size={10} style={{ color: '#6B5B3E' }} /></button>
                              {towerSettingsOpen && (
                                <button onClick={() => requestDelete(() => removeTowerEntry(t.id), `'${t.book_title}'을(를) 책장에서 없앨까요?`)} className="p-0.5" aria-label="책탑에서 제거"><X size={10} style={{ color: '#6B5B3E' }} /></button>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
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
              <div className="flex justify-center"><PrimaryBtn onClick={startSession} icon={QrCode}>오늘 출결 시작</PrimaryBtn></div>
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
  // 주차 헤더 라벨("1주차(8.31-9.3)")과, 그 아래 도트 행이 세로로 정확히 정렬되도록 두 행이 공유할 컬럼 폭을 미리 계산
  // 달력 기준 주차가 아니라, 화면에 보이는 순서대로 1,2,3...으로 매겨서 번호가 겹치는 일이 없게 함
  const weekLabels = weekChunkRanges.map(([start, end], wi) => {
    const chunk = monthDayList.slice(start, end);
    const d1 = chunk[0].date, d2 = chunk[chunk.length - 1].date;
    const month1 = parseInt(d1.slice(5, 7), 10), day1 = parseInt(d1.slice(8, 10), 10);
    const month2 = parseInt(d2.slice(5, 7), 10), day2 = parseInt(d2.slice(8, 10), 10);
    const range = month1 !== month2 ? `${month1}.${day1}-${month2}.${day2}` : (day1 === day2 ? `${month1}.${day1}` : `${month1}.${day1}-${day2}`);
    return { main: `${wi + 1}주차`, sub: `(${range})` };
  });
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
                const dayPenaltyFolks = penaltyEntries.filter((e) => e.completion?.performed_date === date); // 회원 누구든 이 날짜에 벌칙을 수행하면 커피 아이콘 표기
                const birthdayFolks = members.filter((m) => m.birthday && mdOf(m.birthday) === date.slice(5, 10));
                const hasBirthday = birthdayFolks.length > 0;
                let borderStyle = isToday ? '1.5px solid rgba(242,238,227,0.55)' : selectedDate === date ? `1.5px solid ${textColor}` : '1px solid transparent';
                if (hasFeast) borderStyle = '1.5px solid rgba(229, 72, 77, 0.65)';
                return (
                  <button key={date} onClick={() => setSelectedDate(date === selectedDate ? null : date)}
                    className="relative aspect-square rounded-lg flex flex-col items-center justify-center text-xs leading-none"
                    style={{ background: bgStyle, color: textColor, border: borderStyle }}>
                    {(hasDiscussion || hasBirthday || dayPenaltyFolks.length > 0) && (
                      <span className="absolute top-0.5 flex items-center gap-0.5">
                        {hasDiscussion && <BookOpen size={8} style={{ color: '#D9C24C' }} />}
                        {dayPenaltyFolks.length > 0 && <Coffee size={8} style={{ color: '#EFC94C', opacity: dayPenaltyFolks.some((e) => e.completion?.confirmed) ? 1 : 0.55 }} />}
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
            <div className="flex flex-nowrap items-center overflow-x-auto mt-3" style={{ gap: 'clamp(3px, 1.6vw, 8px)' }}>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: dayTypeMeta('독서일').color }} /> 독서일</span>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: dayTypeMeta('토론회').color }} /> 토론회</span>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: 'transparent', border: '1.5px solid rgba(229, 72, 77, 0.65)' }} /> 회식일</span>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: dayTypeMeta('휴무일').color }} /> 휴무일</span>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: WEEKEND_TEXT }} /> 금·토·일(제외)</span>
            </div>
            <div className="flex flex-nowrap items-center mt-1.5 overflow-x-auto" style={{ gap: 'clamp(3px, 1.6vw, 8px)' }}>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}>
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
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#F0A87C' }}><path d="M2 15V10L6 6H21V15Z" /><rect x="9" y="8" width="4" height="3.5" /><rect x="15" y="8" width="4" height="3.5" /><line x1="2" y1="15" x2="21" y2="15" /></svg> 출장</span>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><Plane size={11} /> 휴가</span>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><Briefcase size={11} style={{ color: '#D9A93A' }} /> 업무</span>
              <span className="inline-flex items-center gap-1 shrink-0" style={{ color: MUTE, fontSize: 'clamp(8.5px, 2.4vw, 11px)' }}><User size={11} style={{ color: '#7FDCCF' }} /> 개인일정</span>
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
                  {(() => {
                    const penaltyFolksOnDate = penaltyEntries.filter((e) => e.completion?.performed_date === selectedDate);
                    if (penaltyFolksOnDate.length === 0) return null;
                    return (
                      <div className="mb-2.5 p-2.5 rounded-xl" style={{ background: NEUTRAL_BG }}>
                        <div className="text-xs mb-1.5 flex items-center gap-1" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}><Coffee size={12} style={{ color: '#EFC94C' }} /> {fmtDate(selectedDate)} 벌칙 수행</div>
                        <div className="flex flex-wrap gap-1.5">
                          {penaltyFolksOnDate.map((e) => (
                            <span key={`${e.weekKey}_${e.member.id}`} className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold" style={{ background: e.completion?.confirmed ? '#12302C' : '#3A2E10', color: e.completion?.confirmed ? '#7FDCCF' : '#EFC94C' }}>
                              <Stamp role={e.member.role} size={16} tilt={0} />{dispName(e.member.name, isLoggedIn)} · {e.completion?.confirmed ? '완료' : '예정'}
                            </span>
                          ))}
                        </div>
                      </div>
                    );
                  })()}
                  {excusedMembers.length > 0 && selectedDate !== todayStr() && (
                    <div className="mb-2.5 p-2.5 rounded-xl" style={{ background: NEUTRAL_BG }}>
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
                    <div className="mb-2.5 p-2.5 rounded-xl" style={{ background: NEUTRAL_BG }}>
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
                  {currentMember && penaltyEntries.some((e) => e.member.id === currentMember.id && !isWeekCompleted(e.weekKey, currentMember.id)) && (
                    <div className="mb-2.5 p-2.5 rounded-xl" style={{ background: NEUTRAL_BG }}>
                      <div className="text-xs mb-1.5 flex items-center gap-1" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}><Coffee size={12} style={{ color: '#EFC94C' }} /> 본인 벌칙 수행일 지정</div>
                      <div className="space-y-1.5">
                        {penaltyEntries.filter((e) => e.member.id === currentMember.id && !isWeekCompleted(e.weekKey, currentMember.id)).map((e) => {
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
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><svg width="11.7" height="11.7" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#F0A87C' }}><path d="M2 15V10L6 6H21V15Z" /><rect x="9" y="8" width="4" height="3.5" /><rect x="15" y="8" width="4" height="3.5" /><line x1="2" y1="15" x2="21" y2="15" /></svg>출장</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><Plane size={9} style={{ color: INK }} />휴가</span>
              <span className="inline-flex items-center gap-1 text-[10px]" style={{ color: MUTE }}><span className="inline-block rounded-full" style={{ width: 8, height: 8, border: `1px solid ${LINE}` }} />결석</span>
            </div>
            {/* 주차 수가 많은 달에도 화면 폭에 맞춰 칸 너비가 균등하게 줄어들 뿐, 가로 스크롤이나 두 줄 줄바꿈이 생기지 않도록 그리드로 구성 */}
            {weekChunkRanges.length > 0 && (
              <div className="grid mb-1.5" style={{ paddingLeft: 34, gridTemplateColumns: `repeat(${weekChunkRanges.length}, minmax(0, 1fr))`, columnGap: 3 }}>
                {weekChunkRanges.map(([start, end], wi) => (
                  <div key={wi} className="min-w-0 flex flex-col items-center" style={{ borderRight: wi < weekChunkRanges.length - 1 ? `1px solid ${ROW_LINE}` : 'none' }}>
                    <span className="truncate max-w-full" style={{ fontSize: 'clamp(8px, 2.6vw, 10px)', lineHeight: '11px', color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{weekLabels[wi].main}</span>
                    <span className="truncate max-w-full" style={{ fontSize: 'clamp(7.5px, 2.3vw, 9px)', lineHeight: '10px', color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{weekLabels[wi].sub}</span>
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
                      {isCurrentMonth && warningMemberIds.has(r.id) && !(penaltyByMember[r.id]?.pending > 0) && <span title="이번 주 열린 독서일을 지금까지 모두 결석 — 벌칙유의" className="inline-flex items-center gap-0.5 rounded-full px-1.5 py-0.5 text-[10px] font-semibold shrink-0" style={{ background: '#3A2213', color: '#F0A87C' }}>⚠️ 벌칙유의</span>}
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      {penaltyByMember[r.id]?.pending > 0 && <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}><Gavel size={10} /> 벌칙 대상 {penaltyByMember[r.id].pending}</span>}
                      <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.present}/{r.denom} · {r.rate}%</span>
                    </div>
                  </div>
                  <div className="grid" style={{ paddingLeft: 34, gridTemplateColumns: `repeat(${weekChunkRanges.length}, minmax(0, 1fr))`, columnGap: 3 }}>
                    {weekChunkRanges.map(([start, end], wi) => (
                      <div key={wi} className="min-w-0 flex items-center justify-center overflow-hidden" style={{ gap: 'clamp(1px, 0.6vw, 3px)', borderRight: wi < weekChunkRanges.length - 1 ? `1px solid ${ROW_LINE}` : 'none' }}>
                          {r.flags.slice(start, end).map((status, i) => {
                            const fromPrevMonth = !monthDayList[start + i].inCurrentMonth;
                            if (status === 'trip') {
                              return <svg key={i} width="11.7" height="11.7" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#F0A87C', opacity: fromPrevMonth ? 0.6 : 1 }}><path d="M2 15V10L6 6H21V15Z" /><rect x="9" y="8" width="4" height="3.5" /><rect x="15" y="8" width="4" height="3.5" /><line x1="2" y1="15" x2="21" y2="15" /></svg>;
                            }
                            if (status === 'vacation') {
                              return <Plane key={i} size={9} style={{ color: INK, opacity: fromPrevMonth ? 0.6 : 1 }} />;
                            }
                            return (
                              <span key={i} className="relative rounded-full shrink-0" style={{
                                width: 'clamp(6px, 2vw, 9px)', height: 'clamp(6px, 2vw, 9px)',
                                background: status === 'full' ? '#7FA8D9' : status === 'half' ? 'linear-gradient(90deg, #7FA8D9 50%, transparent 50%)' : status === 'holiday' ? '#E0958C' : 'transparent',
                                border: status === 'full' || status === 'holiday' ? 'none' : `1.5px solid ${LINE}`,
                              }} title={(status === 'holiday' ? '휴무일' : '') + (fromPrevMonth ? ' (이전 달, 참고용)' : '')}>
                                {fromPrevMonth && <span className="absolute inset-0 rounded-full pointer-events-none" style={{ background: 'repeating-linear-gradient(45deg, rgba(30,28,22,0.55) 0px, rgba(30,28,22,0.55) 1px, transparent 1px, transparent 2.5px)' }} />}
                              </span>
                            );
                          })}
                      </div>
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
                  <p className="text-[10px] mt-2" style={{ color: MUTE }}>※ 벌칙 확정된 주 금요일까지 수행 예정일을 지정하고, 실제로 수행한 뒤엔 "완료로 확정"을 눌러야 완료 처리돼요.</p>
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
          <button onClick={downloadExcel} className="flex-1 min-w-0 flex items-center justify-center gap-1.5 rounded-xl py-2.5 px-2 text-sm font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
            <Download size={15} /> 명단 다운로드
          </button>
          <label className="flex-1 min-w-0 flex items-center justify-center gap-1.5 rounded-xl py-2.5 px-2 text-sm font-semibold cursor-pointer" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
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
              <div className="flex flex-nowrap items-center justify-between px-4 py-3">
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
                <div className="flex items-center gap-1 shrink-0 pl-2 ml-1" style={{ borderLeft: `1px solid ${ROW_LINE}` }}>
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

function AdminScreen({ members, sessions, checkins, penaltyRule, setPenaltyRule, penaltyCompletions, reload, calendarDays, absenceExcuses, requestDelete, currentMember, weatherOverride, setWeatherOverride }) {
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
        <div className="flex items-center gap-1.5 text-sm font-semibold mb-2" style={{ color: INK }}><Settings2 size={16} style={{ color: '#7FA8D9' }} /> 화면 효과 설정</div>
        <p className="text-xs mb-2" style={{ color: MUTE }}>기본은 자동(날씨 연동)이고, 생일자가 있는 날은 항상 꽃가루가 우선이에요. 수동으로 고르면 그 효과가 고정으로 나와요.</p>
        <div className="flex flex-wrap gap-1.5">
          {[
            { key: 'auto', label: '자동(날씨 연동)' },
            { key: 'petal', label: '🌸 꽃가루' },
            { key: 'rain', label: '🌧️ 비' },
            { key: 'snow', label: '❄️ 눈' },
            { key: 'off', label: '끄기' },
          ].map((opt) => (
            <button key={opt.key} onClick={() => setWeatherOverride(opt.key)}
              className="rounded-full border px-3 py-1.5 text-xs"
              style={{ borderColor: weatherOverride === opt.key ? '#7FA8D9' : LINE, background: weatherOverride === opt.key ? '#1B3A5C' : 'transparent', color: weatherOverride === opt.key ? '#7FA8D9' : MUTE }}>
              {opt.label}
            </button>
          ))}
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
