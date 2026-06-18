export interface TranscriptEntry {
  time: string;
  text: string;
}

export interface MeetingSummary {
  주요논의: string[];
  결정사항: string[];
  검토필요사항: string[];
  액션아이템: string[];
}

export interface Meeting {
  id: string;
  title: string;
  siteName: string;
  date: string;
  participants: string[];
  transcript: TranscriptEntry[];
  summary: MeetingSummary | null;
  duration: string;
  createdAt: number;
}

export interface SharedMeeting extends Meeting {
  shareId: string;
  fromEmail: string;
  sharedAt: number;
}
