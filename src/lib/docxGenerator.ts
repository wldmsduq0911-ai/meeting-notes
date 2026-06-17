import {
  Document, Paragraph, TextRun, HeadingLevel,
  AlignmentType, Packer, BorderStyle,
} from 'docx';
import type { Meeting } from '@/types/meeting';

export async function generateDocx(meeting: Meeting): Promise<Blob> {
  const children: Paragraph[] = [];

  // 제목
  children.push(
    new Paragraph({
      text: meeting.title,
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
    }),
    new Paragraph({
      children: [new TextRun({ text: `현장명: ${meeting.siteName || '미입력'}`, size: 22 })],
    }),
    new Paragraph({
      children: [new TextRun({ text: `일시: ${meeting.date}`, size: 22 })],
    }),
    new Paragraph({
      children: [new TextRun({ text: `참석자: ${meeting.participants.length > 0 ? meeting.participants.join(', ') : '미입력'}`, size: 22 })],
    }),
    new Paragraph({
      children: [new TextRun({ text: `회의 시간: ${meeting.duration}`, size: 22 })],
    }),
    new Paragraph({ text: '' }),
  );

  // 1. AI 요약본
  if (meeting.summary) {
    children.push(
      new Paragraph({ text: '1. AI 요약본', heading: HeadingLevel.HEADING_2 }),
      new Paragraph({
        children: [
          new TextRun({
            text: '※ AI가 자동 생성한 요약입니다. 내용 누락·오류가 있을 수 있으니 원문을 함께 확인하세요.',
            italics: true,
            color: '888888',
            size: 18,
          }),
        ],
      }),
      new Paragraph({ text: '' }),
    );

    const sections = [
      { key: '주요논의' as const,     label: '■ 주요 논의' },
      { key: '결정사항' as const,     label: '■ 결정 사항' },
      { key: '검토필요사항' as const, label: '■ 검토 필요 사항' },
      { key: '액션아이템' as const,   label: '■ 액션 아이템' },
    ];

    for (const { key, label } of sections) {
      const items = meeting.summary[key] ?? [];
      children.push(
        new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 22 })] }),
        ...(items.length > 0
          ? items.map(item => new Paragraph({ children: [new TextRun({ text: `  • ${item}`, size: 20 })] }))
          : [new Paragraph({ children: [new TextRun({ text: '  • (없음)', size: 20, color: '999999' })] })]),
        new Paragraph({ text: '' }),
      );
    }
  }

  // 2. 전체 대화 내용
  children.push(
    new Paragraph({ text: '2. 대화 순서별 전체 내용', heading: HeadingLevel.HEADING_2 }),
    new Paragraph({
      children: [
        new TextRun({ text: '※ 음성 인식 자동 변환 내용입니다. 일부 오인식이 있을 수 있습니다.', italics: true, color: '888888', size: 18 }),
      ],
    }),
    new Paragraph({ text: '' }),
    ...(meeting.transcript.length > 0
      ? meeting.transcript.map(e =>
          new Paragraph({
            children: [
              new TextRun({ text: `[${e.time}]  `, bold: true, size: 20, color: '1d4ed8' }),
              new TextRun({ text: e.text, size: 20 }),
            ],
          }),
        )
      : [new Paragraph({ children: [new TextRun({ text: '(녹음된 내용 없음)', size: 20, color: '999999' })] })]),
  );

  const doc = new Document({ sections: [{ children }] });
  return Packer.toBlob(doc);
}
