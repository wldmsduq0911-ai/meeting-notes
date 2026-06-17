import {
  Document, Paragraph, TextRun, AlignmentType, Packer, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType,
  Header, Footer, PageNumber,
} from 'docx';
import type { Meeting, MeetingSummary } from '@/types/meeting';

// 색상 (PDF 템플릿 기준)
const NAVY  = '2E4A6F';  // 진한 네이비 (섹션 제목, 표 헤더)
const LBLUE = 'C9DFEF';  // 연한 파랑 (라벨 셀)
const WHITE = 'FFFFFF';
const FONT  = 'Malgun Gothic';

// 페이지 너비 DXA (A4, 좌우여백 2.54cm 기준, 9072 DXA)
const PW = 9072;

// ── 헬퍼 ──────────────────────────────────────────────

function run(text: string, opts: { bold?: boolean; color?: string; size?: number; italic?: boolean } = {}) {
  return new TextRun({ text, font: FONT, size: opts.size ?? 20, bold: opts.bold, color: opts.color, italics: opts.italic });
}

function sectionHead(text: string): Paragraph {
  return new Paragraph({
    children: [run(text, { bold: true, color: NAVY, size: 24 })],
    spacing: { before: 300, after: 120 },
  });
}

function bullet(text: string): Paragraph {
  return new Paragraph({
    children: [run(`  • ${text}`)],
    spacing: { after: 60 },
  });
}

// 셀 배경색 설정 (ShadingType.CLEAR = fill만 사용, 가장 안정적)
function shading(fill: string) {
  return { type: ShadingType.CLEAR, color: 'auto', fill };
}

// 라벨 셀 (연파랑 배경, 굵은 글씨, 가운데 정렬)
function labelCell(text: string, widthDxa: number): TableCell {
  return new TableCell({
    children: [new Paragraph({
      children: [run(text, { bold: true })],
      alignment: AlignmentType.CENTER,
    })],
    shading: shading(LBLUE),
    width: { size: widthDxa, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
  });
}

// 헤더 셀 (네이비 배경, 흰 글씨, 가운데 정렬)
function headerCell(text: string, widthDxa: number): TableCell {
  return new TableCell({
    children: [new Paragraph({
      children: [run(text, { bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
    })],
    shading: shading(NAVY),
    width: { size: widthDxa, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
  });
}

// 데이터 셀 (흰 배경)
function dataCell(text: string, widthDxa: number, colSpan?: number): TableCell {
  return new TableCell({
    children: [new Paragraph({
      children: [run(text)],
    })],
    width: { size: widthDxa, type: WidthType.DXA },
    columnSpan: colSpan,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
  });
}

// 번호 셀
function numCell(n: number, widthDxa: number): TableCell {
  return new TableCell({
    children: [new Paragraph({
      children: [run(String(n))],
      alignment: AlignmentType.CENTER,
    })],
    width: { size: widthDxa, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 80, right: 80 },
  });
}

// 액션아이템 파서: "담당자: 할일 → 기한" 형식 → 분리
function parseAction(text: string): { task: string; person: string; deadline: string } {
  const m = text.match(/^([^:：]{1,10})[：:]\s*(.+?)(?:\s*[→>]\s*(.+))?$/);
  if (m) {
    return {
      person: m[1].trim(),
      task: m[2].replace(/\s*[→>].+$/, '').trim(),
      deadline: m[3]?.trim() ?? '',
    };
  }
  return { task: text, person: '', deadline: '' };
}

// ── 메인 ──────────────────────────────────────────────

export async function generateDocx(meeting: Meeting): Promise<Blob> {
  const S          = meeting.summary;
  const 주요논의   = S?.주요논의    ?? [];
  const 결정사항   = S?.결정사항    ?? [];
  const 검토필요   = S?.검토필요사항 ?? [];
  const 액션아이템 = S?.액션아이템  ?? [];

  // ── 기본 정보 테이블
  // 4열 구조: [라벨(1200) | 내용(3900) | 라벨(900) | 내용(3072)]
  const L1 = 1200, C1 = 3900, L2 = 900, C2 = PW - L1 - C1 - L2; // 3072
  const infoTable = new Table({
    width: { size: PW, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [labelCell('회의명', L1), dataCell(meeting.title, C1 + L2 + C2, 3)] }),
      new TableRow({ children: [labelCell('일시', L1), dataCell(meeting.date, C1), labelCell('장소', L2), dataCell(meeting.siteName || '', C2)] }),
      new TableRow({ children: [labelCell('참석자', L1), dataCell(meeting.participants.join(', '), C1 + L2 + C2, 3)] }),
      new TableRow({ children: [labelCell('작성자', L1), dataCell('', C1), labelCell('부서', L2), dataCell('', C2)] }),
    ],
  });

  // ── 주요 안건 테이블
  // 3열: [No(630) | 논의내용(5292) | 결정/비고(3150)]
  const AN = 630, AC = 5292, AR = PW - AN - AC; // 3150
  const agendaCount = Math.max(주요논의.length, 결정사항.length, 3);
  const agendaTable = new Table({
    width: { size: PW, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [headerCell('No.', AN), headerCell('논의 내용', AC), headerCell('결정 사항 / 비고', AR)] }),
      ...Array.from({ length: agendaCount }, (_, i) =>
        new TableRow({ children: [numCell(i + 1, AN), dataCell(주요논의[i] ?? '', AC), dataCell(결정사항[i] ?? '', AR)] })
      ),
    ],
  });

  // ── Action Items 테이블
  // 3열: [할일(4968) | 담당자(2268) | 기한(1836)]
  const AT = 4968, AP = 2268, AD = PW - AT - AP; // 1836
  const actionCount = Math.max(액션아이템.length, 3);
  const actionTable = new Table({
    width: { size: PW, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [headerCell('할 일', AT), headerCell('담당자', AP), headerCell('기한', AD)] }),
      ...Array.from({ length: actionCount }, (_, i) => {
        const p = 액션아이템[i] ? parseAction(액션아이템[i]) : { task: '', person: '', deadline: '' };
        return new TableRow({ children: [dataCell(p.task, AT), dataCell(p.person, AP), dataCell(p.deadline, AD)] });
      }),
    ],
  });

  // ── 다음 회의 일정 테이블
  const nextTable = new Table({
    width: { size: PW, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [labelCell('일시', L1), dataCell('', C1), labelCell('장소', L2), dataCell('', C2)] }),
      new TableRow({ children: [labelCell('안건', L1), dataCell('', C1 + L2 + C2, 3)] }),
    ],
  });

  // ── 헤더 / 푸터
  const pageHeader = new Header({
    children: [new Paragraph({
      children: [run('회의록 (Minutes of Meeting)', { size: 16, color: '666666' })],
      alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
    })],
  });

  const pageFooter = new Footer({
    children: [new Paragraph({
      children: [
        run('- ', { size: 16, color: '666666' }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, font: FONT, color: '666666' }),
        run(' -', { size: 16, color: '666666' }),
      ],
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'cccccc', space: 4 } },
    })],
  });

  // ── 전체 대화 내용 (별지)
  const transcriptParagraphs: Paragraph[] = [
    new Paragraph({
      children: [run('[ 대화 순서별 전체 내용 ]', { bold: true, color: NAVY, size: 26 })],
      alignment: AlignmentType.CENTER,
      pageBreakBefore: true,
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      children: [run('※ 음성 인식 자동 변환 내용입니다. 일부 오인식이 있을 수 있습니다.', { italic: true, color: '888888', size: 18 })],
      spacing: { after: 200 },
    }),
    ...(meeting.transcript.length > 0
      ? meeting.transcript.map(e => new Paragraph({
          children: [run(`[${e.time}]  `, { bold: true, color: '1d4ed8' }), run(e.text)],
          spacing: { after: 60 },
        }))
      : [new Paragraph({ children: [run('(녹음된 내용 없음)', { color: '999999' })] })]),
  ];

  const doc = new Document({
    sections: [{
      headers: { default: pageHeader },
      footers: { default: pageFooter },
      children: [
        // 제목
        new Paragraph({
          children: [run('회  의  록', { bold: true, color: NAVY, size: 52 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 400 },
        }),

        // 기본 정보 표
        infoTable,

        // AI 면책
        new Paragraph({
          children: [run('※ 아래 내용은 AI가 자동 생성한 요약입니다. 내용 누락·오류가 있을 수 있으니 원문을 함께 확인하세요.', { italic: true, color: '888888', size: 18 })],
          spacing: { before: 240, after: 80 },
        }),

        // 1. 회의 목적
        sectionHead('1. 회의 목적'),
        ...(주요논의.length > 0 ? 주요논의.map(bullet) : [bullet('')]),

        // 2. 주요 안건 및 논의 내용
        sectionHead('2. 주요 안건 및 논의 내용'),
        agendaTable,

        // 3. 결정 사항
        sectionHead('3. 결정 사항'),
        ...(결정사항.length > 0 ? 결정사항.map(bullet) : [bullet(''), bullet(''), bullet('')]),

        // 4. 검토 필요 사항
        sectionHead('4. 검토 필요 사항'),
        ...(검토필요.length > 0 ? 검토필요.map(bullet) : [bullet(''), bullet(''), bullet('')]),

        // 5. Action Items
        sectionHead('5. Action Items (후속 조치)'),
        actionTable,

        // 6. 다음 회의 일정
        sectionHead('6. 다음 회의 일정'),
        nextTable,

        // 서명란
        new Paragraph({
          children: [run('작성자:                              (서명)               확인자:                              (서명)')],
          alignment: AlignmentType.CENTER,
          spacing: { before: 600, after: 0 },
        }),

        // 별지: 전체 대화
        ...transcriptParagraphs,
      ],
    }],
  });

  return Packer.toBlob(doc);
}
