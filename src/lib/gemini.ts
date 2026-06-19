import { GoogleGenerativeAI } from '@google/generative-ai';
import type { MeetingSummary, TranscriptEntry } from '@/types/meeting';

const DOMAIN_HINT = '건설/건축 현장 회의입니다. 음성인식 오류가 있더라도 문맥으로 추론해 처리하세요.';

const JSON_RULES = `규칙:
- 각 항목이 없으면 빈 배열 []로 표시
- 액션아이템은 "[담당자] 내용" 형식 (담당자 모르면 "[미정]" 사용)
- JSON 외 다른 텍스트 절대 포함 금지`;

// 15 MB 미만: 인라인 base64, 이상: Files API 업로드
const INLINE_LIMIT = 15 * 1024 * 1024;

async function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// Gemini Files API 파일 삭제 (fileUri = 업로드 응답의 data.file.uri)
async function deleteGeminiFile(fileUri: string, apiKey: string): Promise<void> {
  try {
    await fetch(`${fileUri}?key=${encodeURIComponent(apiKey)}`, { method: 'DELETE' });
  } catch { /* 삭제 실패는 무시 — 48시간 후 자동 삭제됨 */ }
}

// Gemini Files API로 오디오 업로드 → fileUri 반환
async function uploadAudioFile(audioBlob: Blob, apiKey: string): Promise<string> {
  const mimeType = audioBlob.type || 'audio/webm';

  // 1단계: 업로드 세션 시작
  const initRes = await fetch(
    `https://generativelanguage.googleapis.com/upload/v1beta/files?uploadType=resumable&key=${encodeURIComponent(apiKey)}`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-Upload-Protocol': 'resumable',
        'X-Goog-Upload-Command': 'start',
        'X-Goog-Upload-Header-Content-Length': String(audioBlob.size),
        'X-Goog-Upload-Header-Content-Type': mimeType,
      },
      body: JSON.stringify({ file: { displayName: 'meeting-audio' } }),
    }
  );
  if (!initRes.ok) throw new Error(`Files API 초기화 실패: ${initRes.status}`);

  const uploadUrl = initRes.headers.get('X-Goog-Upload-URL');
  if (!uploadUrl) throw new Error('업로드 URL 없음');

  // 2단계: 파일 전송
  const uploadRes = await fetch(uploadUrl, {
    method: 'POST',
    headers: {
      'Content-Length': String(audioBlob.size),
      'X-Goog-Upload-Offset': '0',
      'X-Goog-Upload-Command': 'upload, finalize',
    },
    body: audioBlob,
  });
  if (!uploadRes.ok) throw new Error(`오디오 업로드 실패: ${uploadRes.status}`);

  const data = await uploadRes.json();
  return data.file.uri as string;
}

// 텍스트 트랜스크립트 → 요약 (오디오 폴백 시)
export async function summarizeMeeting(
  transcript: TranscriptEntry[],
  participants: string[],
  apiKey: string,
): Promise<MeetingSummary> {
  const genAI = new GoogleGenerativeAI(apiKey);
  const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

  const text = transcript.map(e => `[${e.time}] ${e.text}`).join('\n');

  const prompt = `${DOMAIN_HINT}
참석자: ${participants.length > 0 ? participants.join(', ') : '미입력'}

[음성인식 원문 — 오류 포함 가능]
${text}

${JSON_RULES}

{
  "주요논의": ["항목1", "항목2"],
  "결정사항": ["항목1"],
  "검토필요사항": ["항목1"],
  "액션아이템": ["[담당자] 내용1"]
}`;

  const result = await model.generateContent(prompt);
  const raw = result.response.text();
  const match = raw.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('Invalid Gemini response');
  return JSON.parse(match[0]) as MeetingSummary;
}

// 오디오 Blob → 전사 + 요약 (파일 크기에 따라 인라인 or Files API 자동 선택)
export async function transcribeAndSummarize(
  audioBlob: Blob,
  participants: string[],
  apiKey: string,
): Promise<{ transcript: TranscriptEntry[]; summary: MeetingSummary }> {
  const genAI = new GoogleGenerativeAI(apiKey);
  const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });
  const mimeType = audioBlob.type || 'audio/webm';

  let uploadedFileUri: string | null = null;
  let audioPart: object;
  if (audioBlob.size < INLINE_LIMIT) {
    audioPart = { inlineData: { mimeType, data: await blobToBase64(audioBlob) } };
  } else {
    uploadedFileUri = await uploadAudioFile(audioBlob, apiKey);
    audioPart = { fileData: { mimeType, fileUri: uploadedFileUri } };
  }

  const prompt = `${DOMAIN_HINT}
참석자: ${participants.length > 0 ? participants.join(', ') : '미입력'}

오디오를 정확히 전사하고 회의 내용을 요약해 주세요.
${JSON_RULES}
- transcript의 time은 오디오 내 발화 시각 "HH:MM:SS" 형식

{
  "transcript": [{"time": "00:00:05", "text": "발화 내용"}],
  "주요논의": ["항목1"],
  "결정사항": ["항목1"],
  "검토필요사항": ["항목1"],
  "액션아이템": ["[담당자] 내용1"]
}`;

  const result = await model.generateContent([audioPart, prompt]);
  const raw = result.response.text();
  if (uploadedFileUri) deleteGeminiFile(uploadedFileUri, apiKey);

  const match = raw.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('Invalid Gemini audio response');

  const parsed = JSON.parse(match[0]);
  return {
    transcript: (parsed.transcript ?? []) as TranscriptEntry[],
    summary: {
      주요논의:    parsed.주요논의    ?? [],
      결정사항:    parsed.결정사항    ?? [],
      검토필요사항: parsed.검토필요사항 ?? [],
      액션아이템:  parsed.액션아이템  ?? [],
    },
  };
}
