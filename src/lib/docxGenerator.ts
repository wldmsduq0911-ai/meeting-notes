import {
  Document, Paragraph, TextRun, AlignmentType, Packer, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType,
  Header, Footer, PageNumber,
} from 'docx';
import type { Meeting, MeetingSummary } from '@/types/meeting';

const NAVY    = '1e3a5f';
const LBLUE   = 'd6e4f0';
const WHITE   = 'FFFFFF';
const FONT    = 'Malgun Gothic';

// ── 헬퍼 ──────────────────────────────────────────────

function run(text: string, opts: { bold?: boolean; color?: string; size?: number; italic?: boolean } = {}) {
  return new TextRun({ text, font: FONT, size: opts.size ?? 20, bold: opts.bold, color: opts.color, italics: opts.italic });
}

function para(children: TextRun[], opts: { align?: typeof AlignmentType[keyof typeof AlignmentType]; before?: number; after?: number; pageBreak?: boolean } = {}) {
  return new Paragraph({
    children,
    alignment: opts.align,
    spacing: { before: opts.before, after: opts.after ?? 60 },
    pageBreakBefore: opts.pageBreak,
  });
}

function sectionHead(text: string) {
  return para([run(text, { bold: true, color: NAVY, size: 24 })], { before: 300, after: 120 });
}

function bullet(text: string) {
  return para([run(`  • ${text}`)], { after: 60 });
}

function emptyBullets(n = 3) {
  return Array.from({ length: n }, () => bullet(''));
}

function makeCell(
  text: string,
  opts: { bg?: string; width?: number; bold?: boolean; center?: boolean; color?: string } = {}
): TableCell {
  const fg = opts.bg === NAVY ? WHITE : (opts.color ?? '000000');
  return new TableCell({
    children: [
      new Paragraph({
        children: [run(text, { bold: opts.bold, color: fg })],
        alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      }),
    ],
    shading: opts.bg ? { type: ShadingType.SOLID, color: 'auto', fill: opts.bg } : undefined,
    width: opts.width !== undefined ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
  });
}

function headerCell(text: string, width?: number) {
  return makeCell(text, { bg: NAVY, bold: true, center: true, width });
}
function labelCell(text: string, width?: number) {
  return makeCell(text, { bg: LBLUE, bold: true, center: true, width });
}
function dataCell(text: string, width?: number) {
  return makeCell(text, { width });
}

// 액션아이템 파서: "담당자: 할일 → 기한" 형식
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
  const S = meeting.summary;
  const 주요논의    = S?.주요논의    ?? [];
  const 결정사항    = S?.결정사항    ?? [];
  const 검토필요    = S?.검토필요사항 ?? [];
  const 액션아이템  = S?.액션아이템  ?? [];

  // ── 기본 정보 테이블 ──
  const infoTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [labelCell('회의명', 13), dataCell(meeting.title, 87)] }),
      new TableRow({ children: [labelCell('일시', 13), dataCell(meeting.date, 47), labelCell('장소', 10), dataCell(meeting.siteName || '', 30)] }),
      new TableRow({ children: [labelCell('참석자', 13), dataCell(meeting.participants.join(', '), 87)] }),
      new TableRow({ children: [labelCell('작성자', 13), dataCell('', 37), labelCell('부서', 10), dataCell('', 40)] }),
    ],
  });

  // ── 주요 안건 테이블 (논의 내용 | 결정 사항/비고) ──
  const agendaCount = Math.max(주요논의.length, 결정사항.length, 3);
  const agendaTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [headerCell('No.', 7), headerCell('논의 내용', 58), headerCell('결정 사항 / 비고', 35)] }),
      ...Array.from({ length: agendaCount }, (_, i) =>
        new TableRow({
          children: [
            makeCell(String(i + 1), { center: true, width: 7 }),
            dataCell(주요논의[i] ?? '', 58),
            dataCell(결정사항[i] ?? '', 35),
          ],
        })
      ),
    ],
  });

  // ── Action Items 테이블 ──
  const actionCount = Math.max(액션아이템.length, 3);
  const actionTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [headerCell('할 일', 55), headerCell('담당자', 25), headerCell('기한', 20)] }),
      ...Array.from({ length: actionCount }, (_, i) => {
        const p = 액션아이템[i] ? parseAction(액션아이템[i]) : { task: '', person: '', deadline: '' };
        return new TableRow({
          children: [dataCell(p.task, 55), dataCell(p.person, 25), dataCell(p.deadline, 20)],
        });
      }),
    ],
  });

  // ── 다음 회의 일정 테이블 ──
  const nextTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [labelCell('일시', 13), dataCell('', 37), labelCell('장소', 10), dataCell('', 40)] }),
      new TableRow({ children: [labelCell('안건', 13), dataCell('', 87)] }),
    ],
  });

  // ── 헤더 / 푸터 ──
  const pageHeader = new Header({
    children: [
      new Paragraph({
        children: [run('회의록 (Minutes of Meeting)', { size: 16, color: '666666' })],
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
      }),
    ],
  });

  const pageFooter = new Footer({
    children: [
      new Paragraph({
        children: [
          run('- ', { size: 16, color: '666666' }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, font: FONT, color: '666666' }),
          run(' -', { size: 16, color: '666666' }),
        ],
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'cccccc', space: 4 } },
      }),
    ],
  });

  // ── 전체 대화 내용 (별지) ──
  const transcriptSection: Paragraph[] = [
    para(
      [run('[ 대화 순서별 전체 내용 ]', { bold: true, color: NAVY, size: 26 })],
      { align: AlignmentType.CENTER, pageBreak: true, before: 0, after: 200 }
    ),
    para(
      [run('※ 음성 인식 자동 변환 내용입니다. 일부 오인식이 있을 수 있습니다.', { italic: true, color: '888888', size: 18 })],
      { after: 200 }
    ),
    ...(meeting.transcript.length > 0
      ? meeting.transcript.map(e =>
          new Paragraph({
            children: [
              run(`[${e.time}]  `, { bold: true, color: '1d4ed8' }),
              run(e.text),
            ],
            spacing: { after: 60 },
          })
        )
      : [para([run('(녹음된 내용 없음)', { color: '999999' })])]),
  ];

  // ── 문서 조립 ──
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

        // 기본 정보
        infoTable,

        // AI 면책
        para(
          [run('※ 아래 내용은 AI가 자동 생성한 요약입니다. 내용 누락·오류가 있을 수 있으니 원문을 함께 확인하세요.', { italic: true, color: '888888', size: 18 })],
          { before: 240, after: 80 }
        ),

        // 1. 회의 목적
        sectionHead('1. 회의 목적'),
        ...(주요논의.length > 0 ? 주요논의.map(bullet) : emptyBullets(1)),

        // 2. 주요 안건 및 논의 내용
        sectionHead('2. 주요 안건 및 논의 내용'),
        agendaTable,

        // 3. 결정 사항
        sectionHead('3. 결정 사항'),
        ...(결정사항.length > 0 ? 결정사항.map(bullet) : emptyBullets()),

        // 4. 검토 필요 사항
        sectionHead('4. 검토 필요 사항'),
        ...(검토필요.length > 0 ? 검토필요.map(bullet) : emptyBullets()),

        // 5. Action Items
        sectionHead('5. Action Items (후속 조치)'),
        actionTable,

        // 6. 다음 회의 일정
        sectionHead('6. 다음 회의 일정'),
        nextTable,

        // 서명란
        new Paragraph({
          children: [run('작성자:                              (서명)               확인자:                              (서명)', { size: 20 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 600, after: 0 },
        }),

        // 별지: 전체 대화
        ...transcriptSection,
      ],
    }],
  });

  return Packer.toBlob(doc);
}
