import { GoogleGenerativeAI } from '@google/generative-ai';
import type { MeetingSummary, TranscriptEntry } from '@/types/meeting';

export async function summarizeMeeting(
  transcript: TranscriptEntry[],
  participants: string[],
  apiKey: string,
): Promise<MeetingSummary> {
  const genAI = new GoogleGenerativeAI(apiKey);
  const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });

  const text = transcript.map(e => `[${e.time}] ${e.text}`).join('\n');

  const prompt = `다음은 회의 전체 내용입니다. 반드시 아래 JSON 형식으로만 응답하세요.
참석자: ${participants.length > 0 ? participants.join(', ') : '미입력'}

회의 내용:
${text}

규칙:
- 각 항목이 없으면 빈 배열 []로 표시
- 액션아이템은 "[담당자] 내용" 형식 (담당자 모르면 "[미정]" 사용)
- JSON 외 다른 텍스트 절대 포함 금지

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
