'use client';
export const dynamic = 'force-dynamic';
import { useState, useRef, useEffect, useCallback } from 'react';
import {
  onAuthStateChanged, signInWithEmailAndPassword,
  createUserWithEmailAndPassword, signOut, type User,
} from 'firebase/auth';
import { auth } from '@/lib/firebase';
import {
  getMeetingsCloud, saveMeetingCloud, deleteMeetingCloud,
  updateMeetingSiteCloud, renameSiteCloud,
} from '@/lib/cloudStorage';
import type { Meeting, MeetingSummary, TranscriptEntry } from '@/types/meeting';
import { summarizeMeeting } from '@/lib/gemini';
import { generateDocx } from '@/lib/docxGenerator';

type AppState = 'setup' | 'recording' | 'summarizing' | 'done';
type Tab = 'new' | 'history';

const SUMMARY_KEYS = ['주요논의', '결정사항', '검토필요사항', '액션아이템'] as const;
const SUMMARY_META: Record<keyof MeetingSummary, { label: string; icon: string; bg: string; text: string; dot: string }> = {
  주요논의:    { label: '주요 논의',   icon: '💬', bg: 'bg-blue-50',    text: 'text-blue-800',    dot: 'bg-blue-400'    },
  결정사항:    { label: '결정 사항',   icon: '✅', bg: 'bg-emerald-50', text: 'text-emerald-800', dot: 'bg-emerald-400' },
  검토필요사항: { label: '검토 필요', icon: '🔎', bg: 'bg-amber-50',   text: 'text-amber-800',   dot: 'bg-amber-400'   },
  액션아이템:  { label: '액션 아이템', icon: '⚡', bg: 'bg-violet-50',  text: 'text-violet-800',  dot: 'bg-violet-400'  },
};

function fmt(sec: number) {
  const h = String(Math.floor(sec / 3600)).padStart(2, '0');
  const m = String(Math.floor((sec % 3600) / 60)).padStart(2, '0');
  const s = String(sec % 60).padStart(2, '0');
  return `${h}:${m}:${s}`;
}

function WaveBars({ active }: { active: boolean }) {
  const heights = [35, 65, 45, 80, 50, 70, 30, 75, 55, 85, 40, 65, 50, 78, 35, 60, 45, 72, 38, 68];
  return (
    <div className="flex items-center gap-[3px]" style={{ height: 44 }}>
      {heights.map((h, i) => (
        <div
          key={i}
          className={`rounded-full transition-colors duration-500 ${active ? 'bg-indigo-500' : 'bg-gray-200'}`}
          style={{
            width: 3,
            height: active ? `${h}%` : '14%',
            animation: active
              ? `wavePulse ${0.55 + (i % 5) * 0.12}s ease-in-out ${(i * 0.045).toFixed(2)}s infinite alternate`
              : 'none',
          }}
        />
      ))}
    </div>
  );
}

// ── 로그인 / 회원가입 화면 ──────────────────────────────────
function AuthScreen() {
  const [mode, setMode] = useState<'login' | 'signup'>('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async () => {
    setError(''); setLoading(true);
    try {
      if (mode === 'signup') {
        await createUserWithEmailAndPassword(auth, email, password);
      } else {
        await signInWithEmailAndPassword(auth, email, password);
      }
    } catch (e: unknown) {
      const code = (e as { code?: string }).code ?? '';
      if (code === 'auth/email-already-in-use') setError('이미 사용 중인 이메일입니다.');
      else if (code === 'auth/weak-password')   setError('비밀번호는 6자 이상이어야 합니다.');
      else if (code === 'auth/invalid-credential' || code === 'auth/wrong-password' || code === 'auth/user-not-found')
        setError('이메일 또는 비밀번호가 올바르지 않습니다.');
      else setError('오류가 발생했습니다. 다시 시도해 주세요.');
    } finally { setLoading(false); }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center px-5">
      <div className="w-full max-w-sm">
        <div className="flex items-center gap-3 mb-8 justify-center">
          <div className="w-12 h-12 bg-indigo-600 rounded-3xl flex items-center justify-center shadow-lg shadow-indigo-200">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/>
              <path d="M19 10v2a7 7 0 0 1-14 0v-2"/>
              <line x1="12" y1="19" x2="12" y2="23"/>
              <line x1="8" y1="23" x2="16" y2="23"/>
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-gray-900">회의록</h1>
        </div>

        <div className="bg-white rounded-3xl p-6 shadow-sm border border-gray-100">
          <div className="flex gap-1 bg-gray-100 rounded-2xl p-1 mb-6">
            {(['login', 'signup'] as const).map(m => (
              <button key={m} onClick={() => { setMode(m); setError(''); }}
                className={`flex-1 py-2.5 text-sm font-semibold rounded-xl transition-all ${
                  mode === m ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-400 hover:text-gray-600'
                }`}>
                {m === 'login' ? '로그인' : '회원가입'}
              </button>
            ))}
          </div>

          <div className="space-y-3">
            <div>
              <label className="block text-xs font-bold text-gray-400 uppercase tracking-widest mb-1.5">이메일</label>
              <input
                type="email" value={email} onChange={e => setEmail(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleSubmit()}
                placeholder="example@email.com"
                className="w-full bg-gray-50 rounded-2xl px-4 py-3.5 text-gray-900 placeholder-gray-400 text-sm outline-none focus:ring-2 focus:ring-indigo-400 transition-all"
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-gray-400 uppercase tracking-widest mb-1.5">비밀번호</label>
              <input
                type="password" value={password} onChange={e => setPassword(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleSubmit()}
                placeholder={mode === 'signup' ? '6자 이상' : '비밀번호 입력'}
                className="w-full bg-gray-50 rounded-2xl px-4 py-3.5 text-gray-900 placeholder-gray-400 text-sm outline-none focus:ring-2 focus:ring-indigo-400 transition-all"
              />
            </div>

            {error && (
              <p className="text-red-500 text-sm bg-red-50 rounded-2xl px-4 py-3">{error}</p>
            )}

            <button onClick={handleSubmit} disabled={loading || !email || !password}
              className="w-full py-4 rounded-3xl font-bold text-sm bg-indigo-600 hover:bg-indigo-700 active:scale-[0.98] disabled:opacity-40 disabled:cursor-not-allowed transition-all text-white shadow-lg shadow-indigo-200 mt-2">
              {loading ? '처리 중…' : mode === 'login' ? '로그인' : '회원가입'}
            </button>
          </div>
        </div>

        <p className="text-center text-xs text-gray-400 mt-5">
          어떤 기기에서든 같은 계정으로 로그인하면<br/>회의록이 자동으로 동기화됩니다.
        </p>
      </div>
    </div>
  );
}

// ── 메인 앱 ────────────────────────────────────────────────
export default function Home() {
  const [user, setUser]               = useState<User | null>(null);
  const [authReady, setAuthReady]     = useState(false);

  const [tab, setTab]                 = useState<Tab>('new');
  const [state, setState]             = useState<AppState>('setup');
  const [title, setTitle]             = useState('');
  const [siteName, setSiteName]       = useState('');
  const [participantInput, setParticipantInput] = useState('');
  const [participants, setParticipants] = useState<string[]>([]);
  const [transcript, setTranscript]   = useState<TranscriptEntry[]>([]);
  const [interim, setInterim]         = useState('');
  const [timer, setTimer]             = useState(0);
  const [meeting, setMeeting]         = useState<Meeting | null>(null);
  const [resultTab, setResultTab]     = useState<'summary' | 'full'>('summary');
  const [history, setHistory]         = useState<Meeting[]>([]);
  const [selected, setSelected]       = useState<Meeting | null>(null);
  const [selectedSite, setSelectedSite] = useState<string | null>(null);
  const [renamingSite, setRenamingSite] = useState<{ old: string; input: string } | null>(null);
  const [movingSite, setMovingSite]   = useState<{ id: string; input: string } | null>(null);
  const [reconnecting, setReconnecting] = useState(false);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const recognitionRef  = useRef<any>(null);
  const isRecRef        = useRef(false);
  const timerRef        = useRef<ReturnType<typeof setInterval> | null>(null);
  const transcriptRef   = useRef<TranscriptEntry[]>([]);
  const timerValRef     = useRef(0);
  const bottomRef       = useRef<HTMLDivElement>(null);

  // 인증 상태 감지
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (u) => {
      setUser(u);
      if (u) {
        const meetings = await getMeetingsCloud(u.uid);
        setHistory(meetings);
      } else {
        setHistory([]);
      }
      setAuthReady(true);
    });
    return unsub;
  }, []);

  useEffect(() => {
    transcriptRef.current = transcript;
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [transcript]);
  useEffect(() => { timerValRef.current = timer; }, [timer]);

  // 앱 백그라운드 → 포그라운드 복귀 시 자동 이어녹음
  useEffect(() => {
    const handleVisibility = () => {
      if (!isRecRef.current) return;
      if (!document.hidden) {
        // 화면 복귀 → 음성인식 재시작
        setReconnecting(true);
        try { recognitionRef.current?.stop(); } catch {}
        setTimeout(() => {
          if (isRecRef.current && recognitionRef.current) {
            try { recognitionRef.current.start(); } catch {}
          }
          setReconnecting(false);
        }, 500);
      }
    };
    document.addEventListener('visibilitychange', handleVisibility);
    return () => document.removeEventListener('visibilitychange', handleVisibility);
  }, []);

  const refreshHistory = useCallback(async () => {
    if (!user) return;
    setHistory(await getMeetingsCloud(user.uid));
  }, [user]);

  const addParticipant = () => {
    const name = participantInput.trim();
    if (name && !participants.includes(name)) setParticipants(p => [...p, name]);
    setParticipantInput('');
  };

  const startMeeting = useCallback(() => {
    if (!title.trim()) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const w = window as any;
    const API = w.SpeechRecognition || w.webkitSpeechRecognition;
    if (!API) {
      alert('이 브라우저는 음성 인식을 지원하지 않습니다.\nChrome 또는 Edge에서 사용해주세요.');
      return;
    }
    setTranscript([]); setInterim(''); setTimer(0);
    isRecRef.current = true;
    setState('recording');

    timerRef.current = setInterval(() => setTimer(t => t + 1), 1000);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rec: any = new API();
    rec.continuous = true;
    rec.interimResults = true;
    rec.lang = 'ko-KR';
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rec.onresult = (e: any) => {
      let final = '', inter = '';
      for (let i = e.resultIndex; i < e.results.length; i++) {
        const t = e.results[i][0].transcript;
        if (e.results[i].isFinal) final += t;
        else inter += t;
      }
      if (final.trim()) {
        const entry: TranscriptEntry = {
          time: new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
          text: final.trim(),
        };
        setTranscript(p => [...p, entry]);
      }
      setInterim(inter);
    };
    rec.onend = () => { if (isRecRef.current) { try { rec.start(); } catch {} } };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rec.onerror = (e: any) => { if (e.error === 'no-speech' || e.error === 'aborted') return; };
    rec.start();
    recognitionRef.current = rec;
  }, [title]);

  const endMeeting = useCallback(async () => {
    if (!user) return;
    isRecRef.current = false;
    if (timerRef.current) clearInterval(timerRef.current);
    recognitionRef.current?.stop();
    setInterim('');
    setState('summarizing');
    const finalTranscript = [...transcriptRef.current];
    const duration = fmt(timerValRef.current);
    const now = new Date();
    const dateStr = now.toLocaleDateString('ko-KR', {
      year: 'numeric', month: 'long', day: 'numeric', weekday: 'long',
    }) + ' ' + now.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
    let summary: MeetingSummary | null = null;
    const apiKey = process.env.NEXT_PUBLIC_GEMINI_API_KEY || '';
    if (finalTranscript.length > 0 && apiKey) {
      try { summary = await summarizeMeeting(finalTranscript, participants, apiKey); }
      catch (e) { console.error('요약 실패:', e); }
    }
    const m: Meeting = {
      id: Date.now().toString(), title: title.trim(), siteName: siteName.trim(),
      date: dateStr, participants, transcript: finalTranscript, summary, duration, createdAt: Date.now(),
    };
    await saveMeetingCloud(user.uid, m);
    await refreshHistory();
    setMeeting(m);
    setResultTab('summary');
    setState('done');
  }, [title, participants, siteName, user, refreshHistory]);

  const downloadWord = async (m: Meeting) => {
    try {
      const blob = await generateDocx(m);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `${m.title}_${m.date}.docx`; a.click();
      URL.revokeObjectURL(url);
    } catch { alert('Word 파일 생성 중 오류가 발생했습니다.'); }
  };

  const resetToSetup = () => {
    setState('setup'); setTitle(''); setSiteName(''); setParticipants([]);
    setTranscript([]); setMeeting(null); setTimer(0);
  };

  const usedSiteNames = [...new Set(history.map(m => m.siteName).filter(Boolean))];
  const siteGroups = usedSiteNames.map(site => ({
    site,
    meetings: history.filter(m => m.siteName === site),
    lastDate: history.filter(m => m.siteName === site).sort((a, b) => b.createdAt - a.createdAt)[0]?.date ?? '',
  })).sort((a, b) => b.meetings[0]?.createdAt - a.meetings[0]?.createdAt);

  const SummaryView = ({ m }: { m: Meeting }) => (
    <div className="space-y-3">
      {m.summary ? (
        <>
          <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-2xl px-4 py-2.5 leading-relaxed">
            AI 자동 생성 요약입니다. 내용 누락·오류가 있을 수 있으니 원문을 함께 확인하세요.
          </p>
          {SUMMARY_KEYS.map(key => {
            const meta = SUMMARY_META[key];
            const items = m.summary![key] ?? [];
            return (
              <div key={key} className={`${meta.bg} rounded-2xl p-4`}>
                <div className="flex items-center gap-2 mb-3">
                  <span className="text-base">{meta.icon}</span>
                  <h3 className={`text-sm font-bold ${meta.text}`}>{meta.label}</h3>
                </div>
                {items.length > 0 ? (
                  <ul className="space-y-2">
                    {items.map((item: string, i: number) => (
                      <li key={i} className="flex gap-2.5 text-sm text-gray-700">
                        <span className={`w-1.5 h-1.5 rounded-full ${meta.dot} shrink-0 mt-1.5`}/>
                        <span className="leading-relaxed">{item}</span>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <p className="text-sm text-gray-400">해당 없음</p>
                )}
              </div>
            );
          })}
        </>
      ) : (
        <div className="bg-gray-50 rounded-2xl p-10 text-center">
          <p className="text-gray-400 text-sm">
            {m.transcript.length === 0
              ? '녹음된 내용이 없어 요약을 생성할 수 없습니다.'
              : '요약 생성에 실패했습니다.'}
          </p>
        </div>
      )}
    </div>
  );

  const FullTranscriptView = ({ m }: { m: Meeting }) => (
    <div className="space-y-3">
      {m.transcript.length > 0 ? m.transcript.map((e, i) => (
        <div key={i} className="flex gap-3">
          <span className="text-indigo-400 text-xs font-mono shrink-0 mt-0.5 pt-px tabular-nums">{e.time}</span>
          <span className="text-gray-700 text-sm leading-relaxed">{e.text}</span>
        </div>
      )) : (
        <div className="py-12 text-center">
          <p className="text-gray-400 text-sm">녹음된 내용이 없습니다.</p>
        </div>
      )}
    </div>
  );

  const MeetingDetail = ({ m, isNew }: { m: Meeting; isNew: boolean }) => (
    <>
      <div className="bg-white rounded-3xl p-5 shadow-sm border border-gray-100 mb-4">
        <h2 className="font-bold text-gray-900 text-lg leading-tight">{m.title}</h2>
        <p className="text-gray-400 text-xs mt-1">{m.date}</p>
        <div className="flex gap-4 text-xs text-gray-400 mt-2.5">
          <span className="flex items-center gap-1">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
            {m.duration}
          </span>
          {m.participants.length > 0 && (
            <span className="flex items-center gap-1">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
              {m.participants.join(', ')}
            </span>
          )}
        </div>
      </div>

      <div className="flex gap-1 bg-gray-100 rounded-2xl p-1 mb-4">
        {(['summary', 'full'] as const).map(t => (
          <button key={t} onClick={() => setResultTab(t)}
            className={`flex-1 py-2.5 text-sm font-semibold rounded-xl transition-all ${
              resultTab === t ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-400 hover:text-gray-600'
            }`}>
            {t === 'summary' ? 'AI 요약' : '전체 내용'}
          </button>
        ))}
      </div>

      <div className="mb-4">
        {resultTab === 'summary' ? <SummaryView m={m}/> : <FullTranscriptView m={m}/>}
      </div>

      <div className="flex gap-3">
        <button onClick={() => downloadWord(m)}
          className="flex-1 py-4 rounded-3xl font-bold text-sm bg-emerald-500 hover:bg-emerald-600 active:scale-[0.98] text-white shadow-lg shadow-emerald-100 transition-all">
          Word 다운로드
        </button>
        {isNew ? (
          <button onClick={resetToSetup}
            className="flex-1 py-4 rounded-3xl font-bold text-sm bg-gray-100 hover:bg-gray-200 text-gray-600 active:scale-[0.98] transition-all">
            새 회의
          </button>
        ) : (
          <button onClick={async () => {
            if (!user) return;
            await deleteMeetingCloud(user.uid, m.id);
            await refreshHistory();
            setSelected(null);
          }}
            className="py-4 px-5 rounded-3xl font-bold text-sm bg-red-50 hover:bg-red-100 text-red-500 active:scale-[0.98] transition-all">
            삭제
          </button>
        )}
      </div>

      {/* 현장 이동 (히스토리에서만) */}
      {!isNew && (
        <div className="mt-3">
          {movingSite?.id === m.id ? (
            <div className="bg-gray-50 rounded-3xl p-4 border border-gray-200">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-2">다른 현장으로 이동</p>
              <input
                list="move-site-list"
                value={movingSite.input}
                onChange={e => setMovingSite({ id: m.id, input: e.target.value })}
                placeholder="현장명 입력 또는 선택"
                onKeyDown={async e => {
                  if (e.key === 'Enter' && movingSite.input.trim() && user) {
                    await updateMeetingSiteCloud(user.uid, m.id, movingSite.input.trim());
                    await refreshHistory();
                    setMovingSite(null);
                    setSelected(null);
                    setSelectedSite(movingSite.input.trim());
                  }
                  if (e.key === 'Escape') setMovingSite(null);
                }}
                className="w-full bg-white rounded-2xl px-4 py-3 text-gray-900 text-sm outline-none focus:ring-2 focus:ring-indigo-400 border border-gray-200 mb-3"
                autoFocus
              />
              <datalist id="move-site-list">
                {usedSiteNames.filter(n => n !== m.siteName).map(n => <option key={n} value={n}/>)}
              </datalist>
              <div className="flex gap-2">
                <button
                  onClick={async () => {
                    if (movingSite.input.trim() && user) {
                      await updateMeetingSiteCloud(user.uid, m.id, movingSite.input.trim());
                      await refreshHistory();
                      setMovingSite(null);
                      setSelected(null);
                      setSelectedSite(movingSite.input.trim());
                    }
                  }}
                  className="flex-1 py-2.5 rounded-2xl bg-indigo-600 text-white text-sm font-bold">
                  이동
                </button>
                <button onClick={() => setMovingSite(null)}
                  className="flex-1 py-2.5 rounded-2xl bg-gray-200 text-gray-600 text-sm font-bold">
                  취소
                </button>
              </div>
            </div>
          ) : (
            <button onClick={() => setMovingSite({ id: m.id, input: m.siteName || '' })}
              className="w-full py-3 rounded-2xl text-sm text-gray-400 hover:text-indigo-600 border border-dashed border-gray-200 hover:border-indigo-300 transition-all flex items-center justify-center gap-1.5">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="15 17 20 12 15 7"/><path d="M4 18v-2a4 4 0 0 1 4-4h12"/>
              </svg>
              다른 현장으로 이동
            </button>
          )}
        </div>
      )}
    </>
  );

  // 인증 초기화 전 로딩
  if (!authReady) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="relative w-16 h-16">
          <div className="absolute inset-0 rounded-full border-4 border-indigo-100"/>
          <div className="absolute inset-0 rounded-full border-4 border-indigo-500 border-t-transparent animate-spin"/>
        </div>
      </div>
    );
  }

  // 미로그인 → 로그인/회원가입 화면
  if (!user) return <AuthScreen />;

  // 로그인 완료 → 메인 앱
  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-lg mx-auto min-h-screen flex flex-col">

        {/* 헤더 */}
        <header className="px-5 pt-6 pb-4 bg-gray-50 sticky top-0 z-10">
          <div className="flex items-center gap-3 mb-5">
            <div className="w-9 h-9 bg-indigo-600 rounded-2xl flex items-center justify-center shadow-md shadow-indigo-200 shrink-0">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/>
                <path d="M19 10v2a7 7 0 0 1-14 0v-2"/>
                <line x1="12" y1="19" x2="12" y2="23"/>
                <line x1="8" y1="23" x2="16" y2="23"/>
              </svg>
            </div>
            <h1 className="text-xl font-bold text-gray-900 flex-1">회의록</h1>
            <div className="flex items-center gap-2">
              <span className="text-xs text-gray-400 truncate max-w-[120px]">{user.email}</span>
              <button onClick={() => signOut(auth)}
                className="text-xs text-gray-400 hover:text-red-500 transition-colors px-2 py-1 rounded-xl hover:bg-red-50">
                로그아웃
              </button>
            </div>
          </div>
          <div className="flex gap-1 bg-gray-200/70 rounded-2xl p-1">
            {(['new', 'history'] as Tab[]).map(t => (
              <button key={t} onClick={() => setTab(t)}
                className={`flex-1 py-2.5 text-sm font-semibold rounded-xl transition-all ${
                  tab === t ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'
                }`}>
                {t === 'new' ? '새 회의' : `히스토리${history.length > 0 ? ` (${history.length})` : ''}`}
              </button>
            ))}
          </div>
        </header>

        <main className="flex-1 overflow-y-auto px-5 pb-10">

          {/* ──── 새 회의 탭 ──── */}
          {tab === 'new' && (
            <>
              {state === 'setup' && (
                <div className="pt-2 space-y-4">
                  <div className="bg-white rounded-3xl p-5 shadow-sm border border-gray-100 space-y-5">
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase tracking-widest mb-2">현장명</label>
                      <input
                        list="site-names"
                        value={siteName} onChange={e => setSiteName(e.target.value)}
                        placeholder="예) 설계PT, 품평회, 견본주택PT"
                        className="w-full bg-gray-50 rounded-2xl px-4 py-3.5 text-gray-900 placeholder-gray-400 text-sm outline-none focus:ring-2 focus:ring-indigo-400 transition-all"
                      />
                      <datalist id="site-names">
                        {usedSiteNames.map(n => <option key={n} value={n}/>)}
                      </datalist>
                    </div>
                    <div className="h-px bg-gray-100"/>
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase tracking-widest mb-2">회의 제목 *</label>
                      <input
                        value={title} onChange={e => setTitle(e.target.value)}
                        onKeyDown={e => e.key === 'Enter' && startMeeting()}
                        placeholder="예) 6월 주간 팀 회의"
                        className="w-full bg-gray-50 rounded-2xl px-4 py-3.5 text-gray-900 placeholder-gray-400 text-sm outline-none focus:ring-2 focus:ring-indigo-400 transition-all"
                      />
                    </div>
                    <div className="h-px bg-gray-100"/>
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase tracking-widest mb-2">
                        참석자 <span className="font-normal normal-case">선택</span>
                      </label>
                      <div className="flex gap-2">
                        <input
                          value={participantInput} onChange={e => setParticipantInput(e.target.value)}
                          onKeyDown={e => e.key === 'Enter' && addParticipant()}
                          placeholder="이름 입력 후 추가"
                          className="flex-1 bg-gray-50 rounded-2xl px-4 py-3 text-gray-900 placeholder-gray-400 text-sm outline-none focus:ring-2 focus:ring-indigo-400 transition-all"
                        />
                        <button onClick={addParticipant}
                          className="bg-indigo-50 hover:bg-indigo-100 text-indigo-600 font-semibold px-5 py-3 rounded-2xl text-sm transition-colors">
                          추가
                        </button>
                      </div>
                      {participants.length > 0 && (
                        <div className="flex flex-wrap gap-2 mt-3">
                          {participants.map(p => (
                            <span key={p} className="flex items-center gap-1.5 bg-indigo-50 text-indigo-700 text-xs font-semibold px-3 py-1.5 rounded-full">
                              {p}
                              <button onClick={() => setParticipants(ps => ps.filter(x => x !== p))}
                                className="text-indigo-400 hover:text-indigo-700 text-sm font-bold leading-none">×</button>
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>

                  <button onClick={startMeeting} disabled={!title.trim()}
                    className="w-full py-4 rounded-3xl font-bold text-base bg-indigo-600 hover:bg-indigo-700 active:scale-[0.98] disabled:opacity-40 disabled:cursor-not-allowed transition-all text-white shadow-lg shadow-indigo-200">
                    회의 시작
                  </button>
                </div>
              )}

              {state === 'recording' && (
                <div className="pt-2 space-y-3 pb-24">
                  <div className="bg-white rounded-3xl p-5 shadow-sm border border-gray-100">
                    <div className="flex items-center justify-between mb-4">
                      <div className="flex items-center gap-2">
                        {reconnecting ? (
                          <>
                            <div className="w-2.5 h-2.5 rounded-full bg-amber-400 animate-pulse"/>
                            <span className="text-sm font-bold text-amber-600">재연결 중…</span>
                          </>
                        ) : (
                          <>
                            <div className="w-2.5 h-2.5 rounded-full bg-red-500 animate-pulse"/>
                            <span className="text-sm font-bold text-gray-700">녹음 중</span>
                          </>
                        )}
                      </div>
                      <span className="font-mono text-2xl font-bold text-gray-900 tabular-nums">{fmt(timer)}</span>
                    </div>
                    <WaveBars active={!reconnecting}/>
                    <p className="text-sm font-medium text-gray-400 mt-3 truncate">{title}</p>
                  </div>

                  <div className="bg-white rounded-3xl p-5 shadow-sm border border-gray-100 overflow-y-auto" style={{ minHeight: 240, maxHeight: 'calc(100vh - 340px)' }}>
                    {transcript.length === 0 && !interim ? (
                      <div className="flex items-center justify-center h-48">
                        <p className="text-gray-400 text-sm">발화를 시작하면 자막이 표시됩니다</p>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        {transcript.map((e, i) => (
                          <div key={i} className="flex gap-3">
                            <span className="text-indigo-400 text-xs font-mono shrink-0 mt-0.5 tabular-nums">{e.time}</span>
                            <span className="text-gray-700 text-sm leading-relaxed">{e.text}</span>
                          </div>
                        ))}
                        {interim && (
                          <div className="flex gap-3 opacity-40">
                            <span className="text-indigo-300 text-xs font-mono shrink-0 mt-0.5">…</span>
                            <span className="text-gray-500 text-sm italic leading-relaxed">{interim}</span>
                          </div>
                        )}
                        <div ref={bottomRef}/>
                      </div>
                    )}
                  </div>

                  {/* 회의 종료 버튼 — 항상 화면 하단 고정 */}
                  <div className="fixed bottom-0 left-0 right-0 px-5 pb-6 pt-3 bg-gradient-to-t from-gray-50 via-gray-50 to-transparent">
                    <div className="max-w-lg mx-auto">
                      <button
                        onClick={endMeeting}
                        className="w-full py-5 rounded-3xl font-bold text-base bg-red-500 active:bg-red-600 text-white shadow-xl shadow-red-200"
                        style={{ touchAction: 'manipulation' }}>
                        회의 종료
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {state === 'summarizing' && (
                <div className="flex flex-col items-center justify-center py-32 gap-6">
                  <div className="relative w-20 h-20">
                    <div className="absolute inset-0 rounded-full border-4 border-indigo-100"/>
                    <div className="absolute inset-0 rounded-full border-4 border-indigo-500 border-t-transparent animate-spin"/>
                    <div className="absolute inset-0 flex items-center justify-center text-xl">✨</div>
                  </div>
                  <div className="text-center">
                    <p className="text-gray-900 text-lg font-bold mb-1">AI 요약 생성 중</p>
                    <p className="text-gray-400 text-sm">잠시만 기다려 주세요</p>
                  </div>
                </div>
              )}

              {state === 'done' && meeting && (
                <div className="pt-2">
                  <MeetingDetail m={meeting} isNew={true}/>
                </div>
              )}
            </>
          )}

          {/* ──── 히스토리 탭 ──── */}
          {tab === 'history' && (
            <div className="pt-2">
              {selected ? (
                <>
                  <button onClick={() => setSelected(null)}
                    className="flex items-center gap-1.5 text-indigo-600 text-sm font-semibold mb-4 hover:opacity-70 transition-opacity">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="15 18 9 12 15 6"/>
                    </svg>
                    {selectedSite || '목록으로'}
                  </button>
                  <MeetingDetail m={selected} isNew={false}/>
                </>
              ) : selectedSite !== null ? (
                <>
                  <button onClick={() => setSelectedSite(null)}
                    className="flex items-center gap-1.5 text-indigo-600 text-sm font-semibold mb-4 hover:opacity-70 transition-opacity">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="15 18 9 12 15 6"/>
                    </svg>
                    전체 현장
                  </button>
                  <div className="bg-indigo-50 rounded-2xl px-4 py-2.5 mb-4 flex items-center gap-2">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#4f46e5" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>
                    </svg>
                    <span className="text-indigo-700 font-bold text-sm">{selectedSite}</span>
                    <span className="text-indigo-400 text-xs ml-auto">{history.filter(m => m.siteName === selectedSite).length}개 회의</span>
                  </div>
                  <div className="space-y-3">
                    {history.filter(m => m.siteName === selectedSite).sort((a, b) => b.createdAt - a.createdAt).map(m => (
                      <button key={m.id} onClick={() => { setSelected(m); setResultTab('summary'); }}
                        className="w-full bg-white rounded-3xl p-5 text-left shadow-sm border border-gray-100 hover:border-indigo-200 hover:shadow-md active:scale-[0.98] transition-all">
                        <div className="flex items-start justify-between gap-3">
                          <p className="font-bold text-gray-900 text-sm leading-tight">{m.title}</p>
                          <span className="text-xs text-gray-400 font-mono shrink-0 tabular-nums">{m.duration}</span>
                        </div>
                        <p className="text-xs text-gray-400 mt-1.5">{m.date}</p>
                        {m.participants.length > 0 && (
                          <div className="flex gap-1.5 mt-2 flex-wrap">
                            {m.participants.map(p => (
                              <span key={p} className="text-xs bg-gray-100 text-gray-500 px-2 py-0.5 rounded-full">{p}</span>
                            ))}
                          </div>
                        )}
                        {m.summary && (
                          <div className="flex gap-1.5 mt-3 flex-wrap">
                            {(Object.keys(SUMMARY_META) as (keyof MeetingSummary)[])
                              .filter(k => (m.summary![k] ?? []).length > 0)
                              .map(k => (
                                <span key={k} className={`text-xs px-2.5 py-0.5 rounded-full font-semibold ${SUMMARY_META[k].bg} ${SUMMARY_META[k].text}`}>
                                  {SUMMARY_META[k].label} {(m.summary![k] ?? []).length}
                                </span>
                              ))}
                          </div>
                        )}
                      </button>
                    ))}
                  </div>
                </>
              ) : history.length === 0 ? (
                <div className="flex flex-col items-center justify-center py-32 gap-3">
                  <div className="w-16 h-16 bg-gray-100 rounded-3xl flex items-center justify-center">
                    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#9CA3AF" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                    </svg>
                  </div>
                  <p className="text-gray-400 text-sm">저장된 회의록이 없습니다</p>
                </div>
              ) : (
                <div className="space-y-3">
                  {siteGroups.length > 0 ? siteGroups.map(({ site, meetings, lastDate }) => (
                    <div key={site}>
                      {renamingSite?.old === site ? (
                        <div className="bg-white rounded-3xl p-5 shadow-sm border border-indigo-200">
                          <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-2">현장명 변경</p>
                          <input
                            list="rename-site-list"
                            value={renamingSite.input}
                            onChange={e => setRenamingSite({ old: site, input: e.target.value })}
                            onKeyDown={async e => {
                              if (e.key === 'Enter' && renamingSite.input.trim() && user) {
                                await renameSiteCloud(user.uid, site, renamingSite.input.trim(), history);
                                await refreshHistory();
                                setRenamingSite(null);
                              }
                              if (e.key === 'Escape') setRenamingSite(null);
                            }}
                            className="w-full bg-gray-50 rounded-2xl px-4 py-3 text-gray-900 text-sm outline-none focus:ring-2 focus:ring-indigo-400 mb-3"
                            autoFocus
                          />
                          <datalist id="rename-site-list">
                            {usedSiteNames.filter(n => n !== site).map(n => <option key={n} value={n}/>)}
                          </datalist>
                          <p className="text-xs text-gray-400 mb-3">다른 현장명 입력 시 해당 현장으로 회의가 통합됩니다.</p>
                          <div className="flex gap-2">
                            <button
                              onClick={async () => {
                                if (renamingSite.input.trim() && user) {
                                  await renameSiteCloud(user.uid, site, renamingSite.input.trim(), history);
                                  await refreshHistory();
                                  setRenamingSite(null);
                                }
                              }}
                              className="flex-1 py-2.5 rounded-2xl bg-indigo-600 text-white text-sm font-bold">
                              변경
                            </button>
                            <button onClick={() => setRenamingSite(null)}
                              className="flex-1 py-2.5 rounded-2xl bg-gray-100 text-gray-600 text-sm font-bold">
                              취소
                            </button>
                          </div>
                        </div>
                      ) : (
                        <div className="bg-white rounded-3xl shadow-sm border border-gray-100 hover:border-indigo-200 hover:shadow-md transition-all overflow-hidden">
                          <button onClick={() => setSelectedSite(site)} className="w-full p-5 text-left">
                            <div className="flex items-center gap-3">
                              <div className="w-10 h-10 bg-indigo-50 rounded-2xl flex items-center justify-center shrink-0">
                                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#4f46e5" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                  <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>
                                </svg>
                              </div>
                              <div className="flex-1 min-w-0">
                                <p className="font-bold text-gray-900 text-sm">{site}</p>
                                <p className="text-xs text-gray-400 mt-0.5 truncate">{lastDate}</p>
                              </div>
                              <div className="text-right shrink-0 mr-2">
                                <p className="text-lg font-bold text-indigo-600">{meetings.length}</p>
                                <p className="text-xs text-gray-400">건</p>
                              </div>
                            </div>
                          </button>
                          <div className="border-t border-gray-100 px-5 py-2.5">
                            <button
                              onClick={() => setRenamingSite({ old: site, input: site })}
                              className="text-xs text-gray-400 hover:text-indigo-600 transition-colors flex items-center gap-1">
                              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                              </svg>
                              현장명 변경 · 통합
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )) : (
                    <div className="space-y-3">
                      {history.map(m => (
                        <button key={m.id} onClick={() => { setSelected(m); setResultTab('summary'); }}
                          className="w-full bg-white rounded-3xl p-5 text-left shadow-sm border border-gray-100 hover:border-indigo-200 hover:shadow-md active:scale-[0.98] transition-all">
                          <div className="flex items-start justify-between gap-3">
                            <p className="font-bold text-gray-900 text-sm leading-tight">{m.title}</p>
                            <span className="text-xs text-gray-400 font-mono shrink-0 tabular-nums">{m.duration}</span>
                          </div>
                          <p className="text-xs text-gray-400 mt-1.5">{m.date}</p>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

        </main>
      </div>
    </div>
  );
}
