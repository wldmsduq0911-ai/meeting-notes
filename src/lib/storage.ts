import type { Meeting } from '@/types/meeting';

const KEY = 'meeting-notes-history';

export function saveMeeting(meeting: Meeting): void {
  const list = getMeetings();
  const idx = list.findIndex(m => m.id === meeting.id);
  if (idx >= 0) list[idx] = meeting;
  else list.unshift(meeting);
  localStorage.setItem(KEY, JSON.stringify(list));
}

export function getMeetings(): Meeting[] {
  if (typeof window === 'undefined') return [];
  try {
    return JSON.parse(localStorage.getItem(KEY) || '[]');
  } catch {
    return [];
  }
}

export function deleteMeeting(id: string): void {
  localStorage.setItem(KEY, JSON.stringify(getMeetings().filter(m => m.id !== id)));
}
