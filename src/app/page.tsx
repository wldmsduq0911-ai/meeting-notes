'use client';
import { useState, useRef, useEffect, useCallback } from 'react';
import type { Meeting, MeetingSummary, TranscriptEntry } from '@/types/meeting';
import { saveMeeting, getMeetings, deleteMeeting } from '@/lib/storage';
import { summarizeMeeting } from '@/lib/gemini';
import { generateDocx } from '@/lib/docxGenerator';

type AppState = 'setup' | 'recording' | 'summarizing' | 'done';
type Tab = 'new' | 'history';

const SUMMARY_LABELS: Record<keyof MeetingSummary, string> = {
  주요논의: '📋 주요 논의',
  결정사항: '✅ 결정 사항',
  검토필요사항: '🔍 검토 필요 사항',
  액션아이템: '⚡ 액션 아이템',
};
const SUMMARY_KEYS = ['주요논의', '결정사항', '검토필요사항', '액션아이템'] as const;

function fmt(sec: number) {
  const h = String(Math.floor(sec / 3600)).padStart(2, '0');
  const m = String(Math.floor((sec % 3600) / 60)).padStart(2, '0');
  const s = String(sec % 60).padStart(2, '0');
  return `${h}:${m}:${s}`;
}

export default function Home() {
  const [tab, setTab] = useState<Tab>('new');
  const [state, setState] = useState<AppState>('setup');

  // 설정
  const [title, setTitle] = useState('');
  const [participantInput, setParticipantInput] = useState('');
  const [participants, setParticipants] = useState<string[]>([]);

  // 녹음
  const [transcript, setTranscript] = useState<TranscriptEntry[]>([]);
  const [interim, setInterim] = useState('');
  const [timer, setTimer] = useState(0);

  // 결과
  const [meeting, setMeeting] = useState<Meeting | null>(null);
  const [resultTab, setResultTab] = useState<'summary' | 'full'>('summary');

  // 히스토리
  const [history, setHistory] = useState<Meeting[]>([]);
  const [selected, setSelected] = useState<Meeting | null>(null);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const recognitionRef = useRef<any>(null);
  const isRecRef = useRef(false);
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const transcriptRef = useRef<TranscriptEntry[]>([]);
  const timerValRef = useRef(0);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => { setHistory(getMeetings()); }, []);

  useEffect(() => {
    transcriptRef.current = transcript;
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [transcript]);

  useEffect(() => {
    timerValRef.current = timer;
  }, [timer]);

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
    rec.onerror = (e: any) => {
      if (e.error === 'no-speech' || e.error === 'aborted') return;
    };
    rec.start();
    recognitionRef.current = rec;
  }, [title]);

  const endMeeting = useCallback(async () => {
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
      try {
        summary = await summarizeMeeting(finalTranscript, participants, apiKey);
      } catch (e) {
        console.error('요약 실패:', e);
      }
    }

    const m: Meeting = {
      id: Date.now().toString(),
      title: title.trim(),
      date: dateStr,
      participants,
      transcript: finalTranscript,
      summary,
      duration,
      createdAt: Date.now(),
    };
    saveMeeting(m);
    setHistory(getMeetings());
    setMeeting(m);
    setResultTab('summary');
    setState('done');
  }, [title, participants]);

  const downloadWord = async (m: Meeting) => {
    try {
      const blob = await generateDocx(m);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${m.title}_${m.date}.docx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch {
      alert('Word 파일 생성 중 오류가 발생했습니다.');
    }
  };

  const resetToSetup = () => {
    setState('setup'); setTitle(''); setParticipants([]); setTranscript([]); setMeeting(null); setTimer(0);
  };

  const SummaryView = ({ m }: { m: Meeting }) => (
    <div className="space-y-4">
      {m.summary ? (
        <>
          <p className="text-xs text-amber-400 bg-amber-400/10 rounded-lg px-3 py-2">
            ※ AI가 자동 생성한 요약입니다. 내용 누락·오류가 있을 수 있으니 원문을 함께 확인하세요.
          </p>
          {SUMMARY_KEYS.map(key => (
            <div key={key} className="bg-slate-800 rounded-xl p-4">
              <h3 className="font-bold text-sm text-slate-300 mb-3">{SUMMARY_LABELS[key]}</h3>
              {(m.summary![key] ?? []).length > 0 ? (
                <ul className="space-y-1.5">
                  {m.summary![key].map((item: string, i: number) => (
                    <li key={i} className="text-sm text-white flex gap-2">
                      <span className="text-blue-400 mt-0.5 shrink-0">•</span>
                      <span>{item}</span>
                    </li>
                  ))}
                </ul>
              ) : (
                <p className="text-sm text-slate-500">없음</p>
              )}
            </div>
          ))}
        </>
      ) : (
        <div className="bg-slate-800 rounded-xl p-6 text-center text-slate-400 text-sm">
          {m.transcript.length === 0
            ? '녹음된 내용이 없어 요약을 생성할 수 없습니다.'
            : 'API 키가 설정되지 않았거나 요약 생성에 실패했습니다.'}
        </div>
      )}
    </div>
  );

  const FullTranscriptView = ({ m }: { m: Meeting }) => (
    <div className="space-y-2">
      {m.transcript.length > 0 ? m.transcript.map((e, i) => (
        <div key={i} className="flex gap-3 text-sm">
          <span className="text-blue-400 shrink-0 font-mono text-xs mt-0.5">{e.time}</span>
          <span className="text-slate-200">{e.text}</span>
        </div>
      )) : (
        <p className="text-slate-500 text-sm text-center py-8">녹음된 내용이 없습니다.</p>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-slate-900 flex flex-col max-w-2xl mx-auto">
      {/* 헤더 */}
      <header className="px-4 pt-safe pt-4 pb-2">
        <div className="flex items-center gap-2 mb-4">
          <span className="text-2xl">🎙</span>
          <h1 className="text-lg font-bold text-white">회의록 자동 작성</h1>
        </div>
        <div className="flex gap-1 bg-slate-800 rounded-xl p-1">
          {(['new', 'history'] as Tab[]).map(t => (
            <button key={t} onClick={() => setTab(t)}
              className={`flex-1 py-2 text-sm font-medium rounded-lg transition-all ${
                tab === t ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
              }`}>
              {t === 'new' ? '새 회의' : `히스토리 (${history.length})`}
            </button>
          ))}
        </div>
      </header>

      <main className="flex-1 overflow-y-auto px-4 pb-8">
        {/* ──────── 새 회의 탭 ──────── */}
        {tab === 'new' && (
          <div className="mt-4 space-y-4">

            {/* SETUP */}
            {state === 'setup' && (
              <>
                <div className="bg-slate-800 rounded-2xl p-5 space-y-4">
                  <div>
                    <label className="text-xs text-slate-400 mb-1.5 block">회의 제목 *</label>
                    <input
                      value={title} onChange={e => setTitle(e.target.value)}
                      onKeyDown={e => e.key === 'Enter' && startMeeting()}
                      placeholder="예) 6월 주간 팀 회의"
                      className="w-full bg-slate-700 rounded-xl px-4 py-3 text-white placeholder-slate-500 text-sm outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  <div>
                    <label className="text-xs text-slate-400 mb-1.5 block">참석자 (선택)</label>
                    <div className="flex gap-2">
                      <input
                        value={participantInput} onChange={e => setParticipantInput(e.target.value)}
                        onKeyDown={e => e.key === 'Enter' && addParticipant()}
                        placeholder="이름 입력 후 추가"
                        className="flex-1 bg-slate-700 rounded-xl px-4 py-3 text-white placeholder-slate-500 text-sm outline-none focus:ring-2 focus:ring-blue-500"
                      />
                      <button onClick={addParticipant}
                        className="bg-slate-700 hover:bg-slate-600 px-4 py-3 rounded-xl text-sm text-white transition-colors">
                        추가
                      </button>
                    </div>
                    {participants.length > 0 && (
                      <div className="flex flex-wrap gap-2 mt-2">
                        {participants.map(p => (
                          <span key={p} className="flex items-center gap-1 bg-blue-600/20 text-blue-300 text-xs px-3 py-1.5 rounded-full">
                            {p}
                            <button onClick={() => setParticipants(ps => ps.filter(x => x !== p))} className="hover:text-white ml-1">×</button>
                          </span>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
                <button onClick={startMeeting} disabled={!title.trim()}
                  className="w-full py-4 rounded-2xl font-bold text-lg bg-blue-600 hover:bg-blue-500 disabled:opacity-40 disabled:cursor-not-allowed transition-all text-white">
                  🎙 회의 시작
                </button>
              </>
            )}

            {/* RECORDING */}
            {state === 'recording' && (
              <>
                {/* 상태 바 */}
                <div className="bg-slate-800 rounded-2xl p-4 flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <span className="pulse-dot w-3 h-3 rounded-full bg-red-500 inline-block"/>
                    <span className="text-sm text-slate-300">녹음 중</span>
                  </div>
                  <span className="font-mono text-xl font-bold text-white">{fmt(timer)}</span>
                </div>

                {/* 회의 정보 */}
                <div className="bg-slate-800/60 rounded-xl px-4 py-2.5 flex items-center gap-3 text-sm">
                  <span className="text-slate-400">📝</span>
                  <span className="text-white font-medium truncate">{title}</span>
                  {participants.length > 0 && (
                    <span className="text-slate-500 shrink-0">| {participants.join(', ')}</span>
                  )}
                </div>

                {/* 실시간 자막 */}
                <div className="bg-slate-800 rounded-2xl p-4 h-80 overflow-y-auto space-y-2">
                  {transcript.length === 0 && !interim && (
                    <p className="text-slate-500 text-sm text-center mt-10">발화를 시작하면 자막이 표시됩니다...</p>
                  )}
                  {transcript.map((e, i) => (
                    <div key={i} className="flex gap-3 text-sm">
                      <span className="text-blue-400 shrink-0 font-mono text-xs mt-0.5">{e.time}</span>
                      <span className="text-slate-200">{e.text}</span>
                    </div>
                  ))}
                  {interim && (
                    <div className="flex gap-3 text-sm opacity-50">
                      <span className="text-blue-400 shrink-0 font-mono text-xs mt-0.5">…</span>
                      <span className="text-slate-300 italic">{interim}</span>
                    </div>
                  )}
                  <div ref={bottomRef}/>
                </div>

                <button onClick={endMeeting}
                  className="w-full py-4 rounded-2xl font-bold text-lg bg-red-600 hover:bg-red-500 transition-all text-white">
                  ⏹ 회의 종료
                </button>
              </>
            )}

            {/* SUMMARIZING */}
            {state === 'summarizing' && (
              <div className="flex flex-col items-center justify-center py-24 gap-4">
                <div className="w-16 h-16 rounded-full border-4 border-blue-500 border-t-transparent animate-spin"/>
                <p className="text-slate-300 text-lg font-medium">AI 요약 생성 중...</p>
                <p className="text-slate-500 text-sm">잠시만 기다려주세요</p>
              </div>
            )}

            {/* DONE */}
            {state === 'done' && meeting && (
              <>
                {/* 회의 정보 헤더 */}
                <div className="bg-slate-800 rounded-2xl p-4 space-y-1">
                  <h2 className="font-bold text-white text-lg">{meeting.title}</h2>
                  <p className="text-slate-400 text-sm">{meeting.date}</p>
                  <div className="flex gap-4 text-xs text-slate-500 mt-1">
                    <span>⏱ {meeting.duration}</span>
                    {meeting.participants.length > 0 && <span>👥 {meeting.participants.join(', ')}</span>}
                  </div>
                </div>

                {/* 탭 */}
                <div className="flex gap-1 bg-slate-800 rounded-xl p-1">
                  {(['summary', 'full'] as const).map(t => (
                    <button key={t} onClick={() => setResultTab(t)}
                      className={`flex-1 py-2 text-sm font-medium rounded-lg transition-all ${
                        resultTab === t ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                      }`}>
                      {t === 'summary' ? '📋 요약본' : '📄 전체 내용'}
                    </button>
                  ))}
                </div>

                <div className="bg-slate-800/50 rounded-2xl p-4">
                  {resultTab === 'summary' ? <SummaryView m={meeting}/> : <FullTranscriptView m={meeting}/>}
                </div>

                {/* 액션 버튼 */}
                <div className="flex gap-3">
                  <button onClick={() => downloadWord(meeting)}
                    className="flex-1 py-3.5 rounded-2xl font-bold bg-green-700 hover:bg-green-600 transition-all text-white text-sm">
                    📥 Word 다운로드
                  </button>
                  <button onClick={resetToSetup}
                    className="flex-1 py-3.5 rounded-2xl font-bold bg-slate-700 hover:bg-slate-600 transition-all text-white text-sm">
                    🎙 새 회의
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {/* ──────── 히스토리 탭 ──────── */}
        {tab === 'history' && (
          <div className="mt-4">
            {selected ? (
              /* 상세 보기 */
              <>
                <button onClick={() => setSelected(null)}
                  className="flex items-center gap-2 text-slate-400 hover:text-white text-sm mb-4 transition-colors">
                  ← 목록으로
                </button>
                <div className="bg-slate-800 rounded-2xl p-4 mb-4 space-y-1">
                  <h2 className="font-bold text-white text-lg">{selected.title}</h2>
                  <p className="text-slate-400 text-sm">{selected.date}</p>
                  <div className="flex gap-4 text-xs text-slate-500">
                    <span>⏱ {selected.duration}</span>
                    {selected.participants.length > 0 && <span>👥 {selected.participants.join(', ')}</span>}
                  </div>
                </div>
                <div className="flex gap-1 bg-slate-800 rounded-xl p-1 mb-4">
                  {(['summary', 'full'] as const).map(t => (
                    <button key={t} onClick={() => setResultTab(t)}
                      className={`flex-1 py-2 text-sm font-medium rounded-lg transition-all ${
                        resultTab === t ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                      }`}>
                      {t === 'summary' ? '📋 요약본' : '📄 전체 내용'}
                    </button>
                  ))}
                </div>
                <div className="bg-slate-800/50 rounded-2xl p-4 mb-4">
                  {resultTab === 'summary' ? <SummaryView m={selected}/> : <FullTranscriptView m={selected}/>}
                </div>
                <div className="flex gap-3">
                  <button onClick={() => downloadWord(selected)}
                    className="flex-1 py-3.5 rounded-2xl font-bold bg-green-700 hover:bg-green-600 text-white text-sm transition-all">
                    📥 Word 다운로드
                  </button>
                  <button onClick={() => { deleteMeeting(selected.id); setHistory(getMeetings()); setSelected(null); }}
                    className="py-3.5 px-5 rounded-2xl font-bold bg-red-900/60 hover:bg-red-800 text-red-300 text-sm transition-all">
                    삭제
                  </button>
                </div>
              </>
            ) : (
              /* 목록 */
              history.length === 0 ? (
                <div className="text-center py-24 text-slate-500">
                  <p className="text-4xl mb-3">📂</p>
                  <p>저장된 회의록이 없습니다.</p>
                </div>
              ) : (
                <div className="space-y-3">
                  {history.map(m => (
                    <button key={m.id} onClick={() => { setSelected(m); setResultTab('summary'); }}
                      className="w-full bg-slate-800 hover:bg-slate-700 rounded-2xl p-4 text-left transition-colors">
                      <div className="flex items-start justify-between gap-2">
                        <p className="font-semibold text-white text-sm leading-tight">{m.title}</p>
                        <span className="text-xs text-slate-500 shrink-0 font-mono">{m.duration}</span>
                      </div>
                      <p className="text-xs text-slate-400 mt-1">{m.date}</p>
                      {m.participants.length > 0 && (
                        <p className="text-xs text-slate-500 mt-0.5">👥 {m.participants.join(', ')}</p>
                      )}
                      {m.summary && (
                        <p className="text-xs text-blue-400 mt-2">
                          AI 요약 · 논의 {m.summary.주요논의.length}건 · 결정 {m.summary.결정사항.length}건 · 액션 {m.summary.액션아이템.length}건
                        </p>
                      )}
                    </button>
                  ))}
                </div>
              )
            )}
          </div>
        )}
      </main>
    </div>
  );
}
