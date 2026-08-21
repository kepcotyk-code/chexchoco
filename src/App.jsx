import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './supabaseClient';
import {
  Crown, Shield, Wallet, User, Plus, Pencil, Trash2, Check, X, Lock, AlertCircle,
  Megaphone, QrCode, BarChart3, Users, Settings2, Download, Upload, ChevronLeft, ChevronRight,
  LogIn, LogOut, Cake, PartyPopper, Archive, Paperclip, FileText, Eye, Pin, Gavel, BookOpen,
  Image as ImageIcon,
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

const DAY_TYPES = [
  { key: '독서일', label: '독서일', color: '#7FDCCF', bg: '#12302C' },
  { key: '휴무일', label: '휴무일', color: '#A39B87', bg: '#2B2820' },
  { key: '토론회', label: '토론회', color: '#EFC94C', bg: '#3A2E10' },
  { key: '회식일', label: '회식일', color: '#F0A87C', bg: '#3A2213' },
];
const dayTypeMeta = (key) => DAY_TYPES.find((d) => d.key === key) || null;

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
const fmtHM = (totalMin) => { const h = Math.floor(totalMin / 60); const m = totalMin % 60; if (h === 0) return `${m}분`; if (m === 0) return `${h}시간`; return `${h}시간 ${m}분`; };
const mdOf = (birthday) => birthday ? birthday.slice(5, 10) : null;
const fmtMD = (md) => { const [m, d] = md.split('-'); return `${parseInt(m, 10)}월 ${parseInt(d, 10)}일`; };
const uid = (prefix) => `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
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

/* ---------- Supabase data layer ---------- */
const TABLES = ['members', 'notices', 'notice_views', 'sessions', 'checkins', 'penalty_completions', 'calendar_days', 'settings', 'photos'];

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
  const [currentUserId, setCurrentUserId] = useState(() => localStorage.getItem('chexchoco-current-user') || null);
  const [tab, setTab] = useState('notice');
  const [showLoginModal, setShowLoginModal] = useState(false);
  const [modalSelectedId, setModalSelectedId] = useState('');
  const [modalPinInput, setModalPinInput] = useState('');
  const [modalPinError, setModalPinError] = useState('');

  const reload = async () => {
    try {
      const data = await fetchAll();
      setMembers(data.members); setNotices(data.notices); setNoticeViews(data.notice_views);
      setSessions(data.sessions); setCheckins(data.checkins); setPenaltyCompletions(data.penalty_completions);
      setCalendarDays(data.calendar_days); setSettings(data.settings); setPhotos(data.photos);
      setError('');
    } catch (e) { setError('데이터를 불러오지 못했어요. 새로고침해 주세요.'); }
    setLoaded(true);
  };
  useEffect(() => { reload(); }, []);

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
  const noManagerExists = !members.some((m) => MANAGE_ROLES.includes(m.role));
  const canManageUsers = noManagerExists || (currentMember ? MANAGE_ROLES.includes(currentMember.role) : false);
  const canManageAttendance = currentMember ? currentMember.role === '간사' : false;

  const sortedMembers = useMemo(() => [...members].sort((a, b) => {
    const ro = roleOrder(a.role) - roleOrder(b.role);
    return ro !== 0 ? ro : a.name.localeCompare(b.name, 'ko');
  }), [members]);

  const TABS = [
    { key: 'notice', label: '공지사항', icon: Megaphone },
    { key: 'gallery', label: '포토로그', icon: ImageIcon },
    { key: 'qr', label: 'QR출결', icon: QrCode },
    { key: 'dashboard', label: '대시보드', icon: BarChart3 },
    { key: 'users', label: '사용자관리', icon: Users },
    ...(canManageAttendance ? [{ key: 'admin', label: '출석관리', icon: Settings2 }] : []),
  ];

  if (!loaded) {
    return <div className="min-h-screen flex items-center justify-center" style={{ background: PAPER_BG }}>
      <div className="text-sm" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>불러오는 중…</div>
    </div>;
  }

  return (
    <div className="min-h-screen" style={{ background: PAPER_BG, fontFamily: "'Inter', sans-serif" }}>
      <div className="max-w-3xl mx-auto px-4 pt-6 pb-24">
        <div className="mb-6 relative">
          <button onClick={() => (currentMember ? logout() : openLogin())}
            className="absolute right-0 top-0 flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold"
            style={{ background: currentMember ? BTN_BG : NEUTRAL_BG, color: currentMember ? BTN_TEXT : NEUTRAL_TEXT }}>
            {currentMember ? <><Stamp role={currentMember.role} size={16} tilt={0} />{currentMember.name}님 · 로그아웃</> : <>로그인</>}
          </button>
          <div className="flex items-center gap-2 mb-1.5">
            <div className="h-px flex-1" style={{ background: LINE }} />
            <span className="text-[11px] tracking-[0.2em] uppercase" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>KEPCO Reading Club</span>
            <div className="h-px flex-1" style={{ background: LINE }} />
          </div>
          <h1 className="text-center text-4xl font-semibold" style={{ fontFamily: "'Fraunces', serif", color: INK }}>책스초코</h1>
        </div>

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

        {tab === 'notice' && <NoticeScreen notices={notices} noticeViews={noticeViews} currentMember={currentMember} canManage={canManageUsers} reload={reload} />}
        {tab === 'gallery' && <GalleryScreen photos={photos} currentMember={currentMember} canManage={canManageUsers} reload={reload} members={members} sessions={sessions} checkins={checkins} />}
        {tab === 'qr' && <QrScreen members={sortedMembers} currentMember={currentMember} sessions={sessions} checkins={checkins} canManage={canManageUsers} canManageAttendance={canManageAttendance} calendarDays={calendarDays} reload={reload} />}
        {tab === 'dashboard' && <DashboardScreen members={sortedMembers} sessions={sessions} checkins={checkins} penaltyRule={penaltyRule} penaltyCompletions={penaltyCompletions} canManage={canManageUsers} calendarDays={calendarDays} reload={reload} />}
        {tab === 'users' && <UsersScreen members={members} sortedMembers={sortedMembers} currentUserId={currentUserId} setIdentity={setIdentity} canManage={canManageUsers} notices={notices} sessions={sessions} checkins={checkins} reload={reload} />}
        {tab === 'admin' && canManageAttendance && <AdminScreen members={sortedMembers} sessions={sessions} checkins={checkins} penaltyRule={penaltyRule} setPenaltyRule={setPenaltyRule} penaltyCompletions={penaltyCompletions} reload={reload} />}
      </div>
    </div>
  );
}

/* ---------------- 공지사항 ---------------- */
const MAX_PDF_BYTES = 3 * 1024 * 1024;

function NoticeScreen({ notices, noticeViews, currentMember, canManage, reload }) {
  const [showForm, setShowForm] = useState(false);
  const [title, setTitle] = useState('');
  const [content, setContent] = useState('');
  const [editingId, setEditingId] = useState(null);
  const [pinned, setPinned] = useState(false);
  const [attachedFile, setAttachedFile] = useState(null);
  const [fileError, setFileError] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [expandedViewsId, setExpandedViewsId] = useState(null);

  const sorted = [...notices].sort((a, b) => {
    if (!!a.pinned !== !!b.pinned) return a.pinned ? -1 : 1;
    return new Date(b.created_at) - new Date(a.created_at);
  });

  const isPdfSignature = (buf) => { const b = new Uint8Array(buf); return b.length >= 4 && b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46; };
  const handleFilePick = (e) => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file) return;
    setFileError('');
    if (file.size > MAX_PDF_BYTES) { setFileError('3MB 이하의 PDF만 첨부할 수 있어요.'); return; }
    const reader = new FileReader();
    reader.onload = (evt) => {
      if (!isPdfSignature(evt.target.result)) { setFileError('PDF 형식이 아니에요.'); return; }
      setAttachedFile(file);
    };
    reader.readAsArrayBuffer(file.slice(0, 8));
  };

  const submit = async () => {
    if (!title.trim() || !content.trim() || submitting) return;
    setSubmitting(true); setFileError('');
    try {
      const id = editingId || uid('n');
      let fileMeta = {};
      if (attachedFile) {
        const path = `${id}.pdf`;
        const { error: upErr } = await supabase.storage.from('notice-files').upload(path, attachedFile, { upsert: true, contentType: 'application/pdf' });
        if (upErr) throw upErr;
        fileMeta = { has_file: true, file_name: attachedFile.name, file_type: 'application/pdf', file_uploaded_at: new Date().toISOString() };
      }
      if (editingId) {
        await updateRow('notices', 'id', editingId, { title: title.trim(), content: content.trim(), pinned, ...fileMeta });
      } else {
        await insertRow('notices', { id, title: title.trim(), content: content.trim(), author_name: currentMember?.name || '익명', created_at: new Date().toISOString(), pinned, ...fileMeta });
      }
      await reload();
      setTitle(''); setContent(''); setShowForm(false); setEditingId(null); setAttachedFile(null); setPinned(false);
    } catch (e) { setFileError('저장에 실패했어요.'); }
    finally { setSubmitting(false); }
  };
  const startEdit = (n) => { setEditingId(n.id); setTitle(n.title); setContent(n.content); setPinned(!!n.pinned); setAttachedFile(null); setFileError(''); setShowForm(true); };
  const remove = async (id) => { await deleteRow('notices', 'id', id); await deleteRow('notice_views', 'notice_id', id); await reload(); };

  const openFile = async (n) => {
    if (currentMember && !noticeViews.some((v) => v.notice_id === n.id && v.member_id === currentMember.id)) {
      await insertRow('notice_views', { id: uid('v'), notice_id: n.id, member_id: currentMember.id, member_name: currentMember.name, viewed_at: new Date().toISOString() });
      reload();
    }
    window.open(publicUrl('notice-files', `${n.id}.pdf`), '_blank');
  };

  return (
    <div className="space-y-3">
      {canManage && (
        showForm ? (
          <Card className="space-y-3">
            <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="제목" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
            <textarea value={content} onChange={(e) => setContent(e.target.value)} placeholder="내용" rows={4} className="w-full rounded-xl border px-3 py-2 text-sm outline-none resize-none" style={inputStyle} />
            <div>
              <label className="inline-flex items-center gap-1.5 rounded-xl py-2 px-3 text-xs font-semibold cursor-pointer" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
                <Paperclip size={13} /> PDF 첨부 (선택, 3MB 이하)
                <input type="file" accept="application/pdf" onChange={handleFilePick} className="hidden" />
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
        return (
          <Card key={n.id} style={n.pinned ? { borderColor: '#EFC94C', borderWidth: 2 } : {}}>
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="flex items-center gap-1.5">
                  {n.pinned && <Pin size={13} style={{ color: '#EFC94C' }} fill="#EFC94C" />}
                  <h3 className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>{n.title}</h3>
                </div>
                <div className="text-xs mt-0.5" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{n.author_name} · {fmtDate(n.created_at)} {fmtTime(n.created_at)}</div>
              </div>
              {canManage && (
                <div className="flex gap-1 shrink-0">
                  <button onClick={() => startEdit(n)} className="p-1.5" style={{ color: MUTE }}><Pencil size={15} /></button>
                  <button onClick={() => remove(n.id)} className="p-1.5" style={{ color: '#F0A87C' }}><Trash2 size={15} /></button>
                </div>
              )}
            </div>
            <p className="text-sm mt-2 whitespace-pre-wrap" style={{ color: NEUTRAL_TEXT }}>{n.content}</p>
            {n.has_file && (
              <div className="mt-3">
                <div className="flex flex-wrap items-center gap-2">
                  <button onClick={() => openFile(n)} className="inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold" style={{ background: NEUTRAL_BG, color: NEUTRAL_TEXT }}>
                    <Paperclip size={13} /> {n.file_name || '첨부파일'} 다운로드
                  </button>
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
function GalleryScreen({ photos, currentMember, canManage, reload, members, sessions, checkins }) {
  const [error, setError] = useState('');
  const [uploading, setUploading] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const [editingDate, setEditingDate] = useState(false);
  const [dateInput, setDateInput] = useState('');
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
  const viewing = sorted.find((p) => p.id === viewingId);
  const saveDate = async () => {
    if (!viewing || !dateInput) return;
    const time = viewing.created_at.slice(11); // 기존 시각(HH:mm:ss.sssZ)은 그대로 유지
    await updateRow('photos', 'id', viewing.id, { created_at: `${dateInput}T${time}` });
    await reload();
    setEditingDate(false);
  };

  const participantsFor = (photo) => {
    const date = photo.created_at.slice(0, 10);
    const session = sessions.find((s) => s.date === date);
    if (!session) return [];
    const ids = checkins.filter((c) => c.session_id === session.id).map((c) => c.member_id);
    return members.filter((m) => ids.includes(m.id));
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
      {sorted.length === 0 ? (
        <Card><p className="text-sm text-center py-6" style={{ color: MUTE }}>아직 올라온 사진이 없어요.</p></Card>
      ) : (
        <div className="grid grid-cols-3 gap-2">
          {sorted.map((p) => (
            <button key={p.id} onClick={() => { setViewingId(p.id); setEditingDate(false); }} className="relative aspect-square rounded-lg overflow-hidden" style={{ background: NEUTRAL_BG }}>
              <img src={publicUrl('photos', p.file_path)} className="w-full h-full object-cover" alt="" loading="lazy" />
              <div className="absolute bottom-1 left-1.5 right-1.5 flex items-center justify-between text-[10px] font-medium" style={{ color: 'rgba(255,255,255,0.75)', textShadow: '0 1px 2px rgba(0,0,0,0.6)', fontFamily: "'IBM Plex Mono', monospace" }}>
                <span>{p.created_at.slice(0, 10)}</span>
                <span className="truncate ml-1">{p.uploader_name}</span>
              </div>
            </button>
          ))}
        </div>
      )}
      {viewing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: 'rgba(0,0,0,0.85)' }} onClick={() => setViewingId(null)}>
          <div className="max-w-full max-h-full flex flex-col items-center" onClick={(e) => e.stopPropagation()}>
            <img src={publicUrl('photos', viewing.file_path)} className="max-w-full max-h-[65vh] rounded-xl" alt="" />
            <div className="flex items-center gap-3 mt-3">
              <span className="text-sm" style={{ color: '#FFFFFF' }}>{viewing.uploader_name} · {fmtDate(viewing.created_at)} {fmtTime(viewing.created_at)}</span>
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => { setEditingDate(!editingDate); setDateInput(viewing.created_at.slice(0, 10)); }} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><Pencil size={15} /></button>
              )}
              {(canManage || viewing.uploader_id === currentMember?.id) && (
                <button onClick={() => removePhoto(viewing)} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#F0A87C' }}><Trash2 size={15} /></button>
              )}
              <button onClick={() => setViewingId(null)} className="p-1.5 rounded-lg" style={{ background: 'rgba(255,255,255,0.15)', color: '#FFFFFF' }}><X size={15} /></button>
            </div>
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
                      {m.name}
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
function QrScreen({ members, currentMember, sessions, checkins, canManage, canManageAttendance, calendarDays, reload }) {
  const today = todayStr();
  const session = sessions.find((s) => s.date === today);

  const startSession = async () => {
    await insertRow('sessions', { id: uid('s'), date: today, created_at: new Date().toISOString() });
    await upsertRow('calendar_days', { date: today, type: '독서일' }, 'date');
    await reload();
  };
  const myCheckin = session ? checkins.find((c) => c.session_id === session.id && c.member_id === currentMember?.id) : null;
  const checkIn = async () => {
    if (!session || !currentMember) return;
    await insertRow('checkins', { id: uid('c'), session_id: session.id, member_id: currentMember.id, check_in_at: new Date().toISOString(), check_out_at: null });
    await reload();
  };
  const checkOut = async () => {
    if (!myCheckin) return;
    await updateRow('checkins', 'id', myCheckin.id, { check_out_at: new Date().toISOString() });
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
  const bulkCheckInAll = async () => {
    if (!session) return;
    const now = new Date().toISOString();
    const additions = members.filter((m) => !getCheckin(m.id)).map((m) => ({ id: uid('c'), session_id: session.id, member_id: m.id, check_in_at: now, check_out_at: null }));
    if (additions.length) { await supabase.from('checkins').insert(additions); await reload(); }
  };
  const bulkCheckOutAll = async () => {
    if (!session) return;
    const now = new Date().toISOString();
    const ids = todaysCheckins.filter((c) => !c.check_out_at).map((c) => c.id);
    if (ids.length) { await supabase.from('checkins').update({ check_out_at: now }).in('id', ids); await reload(); }
  };

  return (
    <div className="space-y-4">
      <Card className="text-center">
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
            <div className="mx-auto rounded-2xl p-6 inline-flex flex-col items-center gap-2 border" style={{ background: NEUTRAL_BG, borderColor: LINE }}>
              <QrCode size={48} color={INK} strokeWidth={1.5} />
              <div className="text-xl font-bold tracking-[0.25em]" style={{ color: INK, fontFamily: "'IBM Plex Mono', monospace" }}>{session.id.slice(-6).toUpperCase()}</div>
            </div>
            <div className="text-xs mt-2" style={{ color: MUTE }}>오늘의 출결 코드</div>
            {currentMember ? (
              <div className="mt-4 flex flex-col items-center gap-2">
                {!myCheckin ? <PrimaryBtn onClick={checkIn} icon={LogIn}>체크인</PrimaryBtn>
                  : !myCheckin.check_out_at ? (
                    <><div className="text-sm" style={{ color: '#7FDCCF' }}>체크인 {fmtTime(myCheckin.check_in_at)}</div><PrimaryBtn onClick={checkOut} icon={LogOut}>체크아웃</PrimaryBtn></>
                  ) : (
                    <div className="text-sm" style={{ color: MUTE }}>
                      체크인 {fmtTime(myCheckin.check_in_at)} → 체크아웃 {fmtTime(myCheckin.check_out_at)}
                      <div className="font-semibold mt-1" style={{ color: durationMin(myCheckin.check_in_at, myCheckin.check_out_at) >= 30 ? '#7FDCCF' : '#F0A87C' }}>
                        {durationMin(myCheckin.check_in_at, myCheckin.check_out_at)}분 · {durationMin(myCheckin.check_in_at, myCheckin.check_out_at) >= 30 ? '출석 인정' : '30분 미만'}
                      </div>
                    </div>
                  )}
              </div>
            ) : <p className="text-sm mt-3" style={{ color: MUTE }}>상단에서 본인을 먼저 선택해 주세요.</p>}
          </div>
        )}
      </Card>

      {session && canManageAttendance && (
        <Card>
          <div className="flex items-center justify-between mb-3">
            <div className="text-sm font-semibold" style={{ color: INK }}>일괄 출결 관리</div>
            <div className="flex gap-2"><GhostBtn onClick={bulkCheckInAll} icon={LogIn}>전체 체크인</GhostBtn><GhostBtn onClick={bulkCheckOutAll} icon={LogOut}>전체 체크아웃</GhostBtn></div>
          </div>
          <div className="space-y-1.5">
            {members.map((m) => {
              const c = getCheckin(m.id); const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
              return (
                <div key={m.id} className="flex items-center justify-between text-sm py-1.5" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                  <div className="flex items-center gap-2 min-w-0"><Stamp role={m.role} size={26} tilt={0} /><span className="truncate" style={{ color: INK }}>{m.name}</span></div>
                  <div className="flex items-center gap-2 shrink-0">
                    {c ? <span className="text-xs" style={{ color: dur === null ? MUTE : dur >= 30 ? '#7FDCCF' : '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtTime(c.check_in_at)}–{fmtTime(c.check_out_at)} {dur !== null && `(${dur}분)`}</span> : <span className="text-xs" style={{ color: MUTE }}>미체크</span>}
                    {!c && <button onClick={() => checkInMember(m.id)} className="p-1.5 rounded-lg" style={{ background: NEUTRAL_BG, color: '#7FDCCF' }}><LogIn size={14} /></button>}
                    {c && !c.check_out_at && <button onClick={() => checkOutMember(m.id)} className="p-1.5 rounded-lg" style={{ background: NEUTRAL_BG, color: '#F0A87C' }}><LogOut size={14} /></button>}
                  </div>
                </div>
              );
            })}
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
                  <div className="flex items-center gap-2"><Stamp role={m?.role || '회원'} size={24} tilt={0} /><span style={{ color: INK }}>{m?.name || '알 수 없음'}</span></div>
                  <span style={{ color: dur === null ? MUTE : dur >= 30 ? '#7FDCCF' : '#F0A87C', fontFamily: "'IBM Plex Mono', monospace" }}>{fmtTime(c.check_in_at)}–{fmtTime(c.check_out_at)} {dur !== null && `(${dur}분)`}</span>
                </div>
              );
            })}
          </div>
        </Card>
      )}
    </div>
  );
}

/* ---------------- 대시보드 ---------------- */
function DashboardScreen({ members, sessions, checkins, penaltyRule, penaltyCompletions, canManage, calendarDays, reload }) {
  const [cursor, setCursor] = useState(new Date());
  const [viewMode, setViewMode] = useState('month');
  const [selectedDate, setSelectedDate] = useState(null);
  const ms = monthStr(cursor);
  const sessionsInMonth = sessions.filter((s) => s.date.startsWith(ms));
  const totalDays = sessionsInMonth.length;
  const shift = (delta) => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + delta, 1));

  const rows = members.map((m) => {
    let present = 0;
    sessionsInMonth.forEach((s) => { const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id); const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null; if (dur !== null && dur >= 30) present++; });
    return { ...m, present, rate: totalDays ? Math.round((present / totalDays) * 100) : 0 };
  });
  const monthSessionIds = new Set(sessionsInMonth.map((s) => s.id));
  const monthReadingRows = members.map((m) => {
    const totalMin = checkins.filter((c) => c.member_id === m.id && monthSessionIds.has(c.session_id)).reduce((sum, c) => { const d = durationMin(c.check_in_at, c.check_out_at); return sum + (d !== null && d > 0 ? d : 0); }, 0);
    return { ...m, totalMin };
  }).sort((a, b) => b.totalMin - a.totalMin);
  const maxMonthReadingMin = Math.max(1, ...monthReadingRows.map((r) => r.totalMin));

  const totalSessions = sessions.length;
  const allTimeRows = members.map((m) => {
    let present = 0;
    sessions.forEach((s) => { const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id); const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null; if (dur !== null && dur >= 30) present++; });
    return { ...m, present, rate: totalSessions ? Math.round((present / totalSessions) * 100) : 0 };
  });
  const totalReadingRows = members.map((m) => {
    const totalMin = checkins.filter((c) => c.member_id === m.id).reduce((sum, c) => { const d = durationMin(c.check_in_at, c.check_out_at); return sum + (d !== null && d > 0 ? d : 0); }, 0);
    return { ...m, totalMin };
  }).sort((a, b) => b.totalMin - a.totalMin);
  const maxReadingMin = Math.max(1, ...totalReadingRows.map((r) => r.totalMin));

  const pastSessions = sessions.filter((s) => s.date < todayStr());
  const penaltyByMember = {};
  members.forEach((m) => { penaltyByMember[m.id] = { pending: 0 }; });
  pastSessions.forEach((s) => members.forEach((m) => {
    const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id);
    const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null;
    if (!(dur !== null && dur >= 30)) { const done = penaltyCompletions.some((p) => p.session_id === s.id && p.member_id === m.id); if (!done) penaltyByMember[m.id].pending += 1; }
  }));

  const todayMd = todayStr().slice(5, 10);
  const isCurrentMonth = ms === monthStr(new Date());
  const monthBirthdays = members.filter((m) => m.birthday && m.birthday.slice(5, 7) === pad(cursor.getMonth() + 1)).sort((a, b) => mdOf(a.birthday).localeCompare(mdOf(b.birthday)));
  const birthdayFolksToday = isCurrentMonth ? monthBirthdays.filter((m) => mdOf(m.birthday) === todayMd) : [];

  const monthGrid = buildMonthGrid(cursor.getFullYear(), cursor.getMonth());
  const getDayType = (date) => calendarDays.find((d) => d.date === date)?.type || null;
  const setDayType = async (date, type) => {
    if (type) await upsertRow('calendar_days', { date, type }, 'date'); else await deleteRow('calendar_days', 'date', date);
    if (type === '독서일') {
      if (!sessions.some((s) => s.date === date)) await insertRow('sessions', { id: uid('s'), date, created_at: new Date().toISOString() });
    } else {
      // 독서일이 아닌 다른 유형으로 바꾸거나 해제할 때, 체크인 기록이 없는 자동 생성 세션은 함께 정리
      const s = sessions.find((ss) => ss.date === date);
      if (s && !checkins.some((c) => c.session_id === s.id)) await deleteRow('sessions', 'id', s.id);
    }
    await reload();
    setSelectedDate(null);
  };

  return (
    <div className="space-y-4">
      {birthdayFolksToday.length > 0 && (
        <Card className="text-center">
          <div className="flex items-center justify-center gap-2 mb-1">
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
            <span className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>오늘은 {birthdayFolksToday.map((m) => m.name).join(', ')}님 생일이에요!</span>
            <PartyPopper size={18} style={{ color: '#EFC94C' }} />
          </div>
          <p className="text-sm" style={{ color: MUTE }}>축하 인사 한마디 건네보는 건 어떨까요 🎂</p>
        </Card>
      )}

      <div className="flex gap-2">
        <button onClick={() => setViewMode('month')} className="flex-1 rounded-xl py-2 text-sm font-semibold" style={{ background: viewMode === 'month' ? BTN_BG : NEUTRAL_BG, color: viewMode === 'month' ? BTN_TEXT : NEUTRAL_TEXT }}>이번 달</button>
        <button onClick={() => setViewMode('all')} className="flex-1 rounded-xl py-2 text-sm font-semibold" style={{ background: viewMode === 'all' ? BTN_BG : NEUTRAL_BG, color: viewMode === 'all' ? BTN_TEXT : NEUTRAL_TEXT }}>전체 기간</button>
      </div>

      {viewMode === 'month' ? (
        <>
          <Card>
            <div className="flex items-center justify-between mb-1">
              <button onClick={() => shift(-1)} className="p-1.5" style={{ color: MUTE }}><ChevronLeft size={18} /></button>
              <div className="font-semibold" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>{cursor.getFullYear()}년 {cursor.getMonth() + 1}월</div>
              <button onClick={() => shift(1)} className="p-1.5" style={{ color: MUTE }}><ChevronRight size={18} /></button>
            </div>
            <div className="text-xs text-center" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>이번 달 출결 {totalDays}회</div>
          </Card>

          <Card>
            <div className="grid grid-cols-7 gap-1 mb-2">
              {['일', '월', '화', '수', '목', '금', '토'].map((w) => <div key={w} className="text-center text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{w}</div>)}
            </div>
            <div className="grid grid-cols-7 gap-1">
              {monthGrid.map((date, idx) => {
                if (!date) return <div key={idx} />;
                const day = parseInt(date.slice(8, 10), 10);
                const type = getDayType(date); const meta = type ? dayTypeMeta(type) : null;
                const isToday = date === todayStr();
                return (
                  <button key={date} onClick={() => canManage && setSelectedDate(date === selectedDate ? null : date)}
                    className="aspect-square rounded-lg flex items-center justify-center text-xs"
                    style={{ background: meta ? meta.bg : NEUTRAL_BG, color: meta ? meta.color : MUTE, border: isToday ? `1.5px solid ${INK}` : selectedDate === date ? `1.5px solid ${meta ? meta.color : MUTE}` : '1px solid transparent' }}>
                    {day}
                  </button>
                );
              })}
            </div>
            <div className="flex flex-wrap gap-2 mt-3">
              {DAY_TYPES.map((t) => <span key={t.key} className="inline-flex items-center gap-1 text-[11px]" style={{ color: MUTE }}><span className="w-2.5 h-2.5 rounded-full" style={{ background: t.color }} /> {t.label}</span>)}
            </div>
            {canManage && selectedDate && (
              <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${ROW_LINE}` }}>
                <div className="text-xs mb-2" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(selectedDate)} 유형 지정</div>
                <div className="flex flex-wrap gap-2">
                  {DAY_TYPES.map((t) => <button key={t.key} onClick={() => setDayType(selectedDate, t.key)} className="rounded-full px-3 py-1.5 text-xs font-semibold" style={{ background: t.bg, color: t.color }}>{t.label}</button>)}
                  <button onClick={() => setDayType(selectedDate, null)} className="rounded-full px-3 py-1.5 text-xs font-semibold" style={{ background: NEUTRAL_BG, color: MUTE }}>지정 해제</button>
                </div>
              </div>
            )}
          </Card>

          {monthBirthdays.length > 0 && (
            <Card>
              <div className="flex items-center gap-1.5 text-sm font-semibold mb-2" style={{ color: INK }}><Cake size={16} style={{ color: '#F0A87C' }} /> 이 달의 생일</div>
              <div className="flex flex-wrap gap-2">
                {monthBirthdays.map((m) => { const t = isCurrentMonth && mdOf(m.birthday) === todayMd; return <span key={m.id} className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold" style={{ background: t ? '#3A2E10' : NEUTRAL_BG, color: t ? '#EFC94C' : NEUTRAL_TEXT }}>{m.name} · {fmtMD(mdOf(m.birthday))}{t && ' 🎉'}</span>; })}
              </div>
            </Card>
          )}

          {penaltyRule && <Card><div className="flex items-center gap-1.5 text-sm font-semibold mb-1" style={{ color: INK }}><Gavel size={16} style={{ color: '#F0A87C' }} /> 벌칙 규정</div><p className="text-sm whitespace-pre-wrap" style={{ color: NEUTRAL_TEXT }}>{penaltyRule}</p></Card>}

          <Card>
            <div className="text-sm font-semibold mb-3" style={{ color: INK }}>이번 달 출석률</div>
            <div className="space-y-3">
              {rows.map((r) => (
                <div key={r.id}>
                  <div className="flex items-center justify-between text-sm mb-1">
                    <div className="flex items-center gap-2"><Stamp role={r.role} size={24} tilt={0} /><span style={{ color: INK }}>{r.name}</span></div>
                    <div className="flex items-center gap-2">
                      {penaltyByMember[r.id]?.pending > 0 && <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}><Gavel size={10} /> 미이행 {penaltyByMember[r.id].pending}</span>}
                      <span style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.present}/{totalDays}회 · {r.rate}%</span>
                    </div>
                  </div>
                  <div className="w-full h-2 rounded-full overflow-hidden" style={{ background: NEUTRAL_BG }}><div className="h-full rounded-full" style={{ width: `${r.rate}%`, background: r.rate >= 80 ? '#7FDCCF' : r.rate >= 50 ? '#EFC94C' : '#F0A87C' }} /></div>
                </div>
              ))}
            </div>
          </Card>

          <Card>
            <div className="flex items-center gap-1.5 text-sm font-semibold mb-3" style={{ color: INK }}><BookOpen size={16} style={{ color: '#EFC94C' }} /> 이번 달 독서시간</div>
            <div className="space-y-3">
              {monthReadingRows.map((r, idx) => (
                <div key={r.id}>
                  <div className="flex items-center justify-between text-sm mb-1">
                    <div className="flex items-center gap-2"><Stamp role={r.role} size={24} tilt={0} /><span style={{ color: INK }}>{r.name}</span>{idx === 0 && r.totalMin > 0 && <span className="text-xs">🏆</span>}</div>
                    <span style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtHM(r.totalMin)}</span>
                  </div>
                  <div className="w-full h-2 rounded-full overflow-hidden" style={{ background: NEUTRAL_BG }}><div className="h-full rounded-full" style={{ width: `${(r.totalMin / maxMonthReadingMin) * 100}%`, background: '#EFC94C' }} /></div>
                </div>
              ))}
            </div>
          </Card>
        </>
      ) : (
        <>
          {penaltyRule && <Card><div className="flex items-center gap-1.5 text-sm font-semibold mb-1" style={{ color: INK }}><Gavel size={16} style={{ color: '#F0A87C' }} /> 벌칙 규정</div><p className="text-sm whitespace-pre-wrap" style={{ color: NEUTRAL_TEXT }}>{penaltyRule}</p></Card>}
          <Card>
            <div className="text-sm font-semibold mb-1" style={{ color: INK }}>전체 누적 출석률</div>
            <div className="text-xs mb-3" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>지금까지 총 출결 {totalSessions}회</div>
            <div className="space-y-3">
              {allTimeRows.map((r) => (
                <div key={r.id}>
                  <div className="flex items-center justify-between text-sm mb-1">
                    <div className="flex items-center gap-2"><Stamp role={r.role} size={24} tilt={0} /><span style={{ color: INK }}>{r.name}</span></div>
                    <span style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{r.present}/{totalSessions}회 · {r.rate}%</span>
                  </div>
                  <div className="w-full h-2 rounded-full overflow-hidden" style={{ background: NEUTRAL_BG }}><div className="h-full rounded-full" style={{ width: `${r.rate}%`, background: r.rate >= 80 ? '#7FDCCF' : r.rate >= 50 ? '#EFC94C' : '#F0A87C' }} /></div>
                </div>
              ))}
            </div>
          </Card>
          <Card>
            <div className="flex items-center gap-1.5 text-sm font-semibold mb-3" style={{ color: INK }}><BookOpen size={16} style={{ color: '#EFC94C' }} /> 누적 독서시간 (전체 기간)</div>
            <div className="space-y-3">
              {totalReadingRows.map((r, idx) => (
                <div key={r.id}>
                  <div className="flex items-center justify-between text-sm mb-1">
                    <div className="flex items-center gap-2"><Stamp role={r.role} size={24} tilt={0} /><span style={{ color: INK }}>{r.name}</span>{idx === 0 && r.totalMin > 0 && <span className="text-xs">🏆</span>}</div>
                    <span style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtHM(r.totalMin)}</span>
                  </div>
                  <div className="w-full h-2 rounded-full overflow-hidden" style={{ background: NEUTRAL_BG }}><div className="h-full rounded-full" style={{ width: `${(r.totalMin / maxReadingMin) * 100}%`, background: '#EFC94C' }} /></div>
                </div>
              ))}
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

/* ---------------- 사용자관리 ---------------- */
function UsersScreen({ members, sortedMembers, currentUserId, setIdentity, canManage, notices, sessions, checkins, reload }) {
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
  const startSelfEdit = (m) => { setEditingId(m.id); setEditMode('self'); setEditBirthday(m.birthday || ''); setEditPin(m.pin || ''); };
  const saveEdit = async () => {
    if (editPin && !/^\d{4}$/.test(editPin)) return;
    if (editMode === 'full') {
      if (!editName.trim()) return;
      await updateRow('members', 'id', editingId, { name: editName.trim(), role: editRole, birthday: editBirthday || null, pin: editPin || null, department: editDept || null, job_type: editJobType || null, joined_at: editJoinedAt || null, book_genre: editGenre || null, note: editNote || null });
    } else {
      await updateRow('members', 'id', editingId, { birthday: editBirthday || null, pin: editPin || null });
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
                {editMode === 'full' && (
                  <>
                    <input value={editName} onChange={(e) => setEditName(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                    <RolePicker value={editRole} onChange={setEditRole} />
                    <div className="grid grid-cols-2 gap-2">
                      <input value={editDept} onChange={(e) => setEditDept(e.target.value)} placeholder="소속" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                      <input value={editJobType} onChange={(e) => setEditJobType(e.target.value)} placeholder="직군" className="rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                    </div>
                    <div><div className="text-xs mb-1" style={{ color: MUTE }}>가입일자</div><input type="date" value={editJoinedAt} onChange={(e) => setEditJoinedAt(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
                    <input value={editGenre} onChange={(e) => setEditGenre(e.target.value)} placeholder="선호 도서 종류 (예: 소설, 자기계발)" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                    <input value={editNote} onChange={(e) => setEditNote(e.target.value)} placeholder="비고" className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} />
                  </>
                )}
                {editMode === 'self' && <p className="text-xs" style={{ color: MUTE }}>본인의 생일과 PIN만 수정할 수 있어요.</p>}
                <div><div className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Cake size={13} /> 생일</div><input type="date" value={editBirthday} onChange={(e) => setEditBirthday(e.target.value)} className="w-full rounded-xl border px-3 py-2 text-sm outline-none" style={inputStyle} /></div>
                <div><div className="text-xs mb-1 flex items-center gap-1" style={{ color: MUTE }}><Lock size={13} /> 본인 확인 PIN (4자리, 선택)</div><input type="password" inputMode="numeric" maxLength={4} value={editPin} onChange={(e) => setEditPin(e.target.value.replace(/\D/g, ''))} className="w-full rounded-xl border px-3 py-2 text-sm tracking-[0.3em] outline-none" style={inputStyle} placeholder="설정 안 함" /></div>
                <div className="flex gap-2 pt-1"><PrimaryBtn onClick={saveEdit} icon={Check}>저장</PrimaryBtn><GhostBtn onClick={() => setEditingId(null)} icon={X}>취소</GhostBtn></div>
              </div>
            ) : (
              <div className="flex items-center justify-between px-4 py-3">
                <div className="flex items-center gap-3 min-w-0 flex-1">
                  <Stamp role={m.role} size={36} tilt={idx % 2 === 0 ? -5 : 4} />
                  <div className="min-w-0 flex-1">
                    <div className="font-semibold truncate" style={{ color: INK }}>{m.name}{m.id === currentUserId && <span className="ml-1.5 text-[11px] font-normal" style={{ color: MUTE }}>(나)</span>}</div>
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <RoleChip role={m.role} />
                      {m.job_type && <span className="hidden sm:inline text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{m.job_type}</span>}
                      {m.birthday && <span className="text-[11px]" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtMD(mdOf(m.birthday))}</span>}
                      {m.pin && <Lock size={11} style={{ color: MUTE }} />}
                    </div>
                    {(m.joined_at || m.book_genre || m.note) && (
                      <div className="hidden sm:flex items-center gap-2 flex-wrap mt-1 text-[11px]" style={{ color: MUTE }}>
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
                  {canManage && <button onClick={() => removeMember(m.id)} className="p-2 rounded-lg" style={{ color: '#F0A87C' }}><Trash2 size={16} /></button>}
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
function AdminScreen({ members, sessions, checkins, penaltyRule, setPenaltyRule, penaltyCompletions, reload }) {
  const [date, setDate] = useState(todayStr());
  const session = sessions.find((s) => s.date === date);
  const dayCheckins = session ? checkins.filter((c) => c.session_id === session.id) : [];
  const [manualMemberId, setManualMemberId] = useState(''); const [manualIn, setManualIn] = useState(''); const [manualOut, setManualOut] = useState('');
  const [editingRule, setEditingRule] = useState(false); const [ruleInput, setRuleInput] = useState(penaltyRule || '');
  const [expandedPenaltyId, setExpandedPenaltyId] = useState(null);

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
  const updateCheckin = async (id, field, timeVal) => { if (!timeVal) return; await updateRow('checkins', 'id', id, { [field]: new Date(`${date}T${timeVal}`).toISOString() }); await reload(); };
  const removeCheckin = async (id) => { await deleteRow('checkins', 'id', id); await reload(); };
  const saveRule = () => { setPenaltyRule(ruleInput.trim()); setEditingRule(false); };
  const pastSessions = sessions.filter((s) => s.date < todayStr()).sort((a, b) => b.date.localeCompare(a.date));
  const isCompleted = (sessionId, memberId) => penaltyCompletions.some((p) => p.session_id === sessionId && p.member_id === memberId);
  const toggleCompletion = async (sessionId, memberId) => {
    if (isCompleted(sessionId, memberId)) {
      const row = penaltyCompletions.find((p) => p.session_id === sessionId && p.member_id === memberId);
      await deleteRow('penalty_completions', 'id', row.id);
    } else await insertRow('penalty_completions', { id: uid('p'), session_id: sessionId, member_id: memberId, completed_at: new Date().toISOString() });
    await reload();
  };
  const penaltyByMember = members.map((m) => {
    const events = [];
    pastSessions.forEach((s) => { const c = checkins.find((ck) => ck.session_id === s.id && ck.member_id === m.id); const dur = c ? durationMin(c.check_in_at, c.check_out_at) : null; if (!(dur !== null && dur >= 30)) events.push({ session_id: s.id, date: s.date }); });
    return { ...m, events, pending: events.filter((e) => !isCompleted(e.session_id, m.id)).length };
  }).filter((m) => m.events.length > 0);

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
                <button onClick={() => removeCheckin(c.id)} className="p-1" style={{ color: '#F0A87C' }}><Trash2 size={14} /></button>
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
        <div className="space-y-2" style={{ borderTop: `1px solid ${ROW_LINE}`, paddingTop: 10 }}>
          {penaltyByMember.length === 0 && <p className="text-sm" style={{ color: MUTE }}>결석 기록이 없어요.</p>}
          {penaltyByMember.map((m) => {
            const expanded = expandedPenaltyId === m.id;
            return (
              <div key={m.id}>
                <button onClick={() => setExpandedPenaltyId(expanded ? null : m.id)} className="w-full flex items-center justify-between py-1.5">
                  <div className="flex items-center gap-2"><Stamp role={m.role} size={26} tilt={0} /><span className="text-sm" style={{ color: INK }}>{m.name}</span></div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs" style={{ color: MUTE, fontFamily: "'IBM Plex Mono', monospace" }}>결석 {m.events.length}회</span>
                    {m.pending > 0 ? <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#3A2213', color: '#F0A87C' }}>미이행 {m.pending}</span> : <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold" style={{ background: '#12302C', color: '#7FDCCF' }}>완료</span>}
                  </div>
                </button>
                {expanded && (
                  <div className="pl-9 pb-2 space-y-1.5">
                    {m.events.map((e) => { const done = isCompleted(e.session_id, m.id); return (
                      <div key={e.session_id} className="flex items-center justify-between text-xs">
                        <span style={{ color: NEUTRAL_TEXT, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtDate(e.date)}</span>
                        <button onClick={() => toggleCompletion(e.session_id, m.id)} className="inline-flex items-center gap-1 rounded-full px-2 py-1 font-semibold" style={{ background: done ? '#12302C' : NEUTRAL_BG, color: done ? '#7FDCCF' : MUTE }}>{done ? <Check size={11} /> : null} {done ? '이행 완료' : '미이행'}</button>
                      </div>
                    ); })}
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
