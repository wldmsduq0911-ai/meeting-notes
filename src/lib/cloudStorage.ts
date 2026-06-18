import { db } from './firebase';
import {
  collection, doc, setDoc, getDocs, deleteDoc,
  query, orderBy, updateDoc, writeBatch, where,
} from 'firebase/firestore';
import type { Meeting, SharedMeeting } from '@/types/meeting';

function meetingsRef(uid: string) {
  return collection(db, 'users', uid, 'meetings');
}

export async function getMeetingsCloud(uid: string): Promise<Meeting[]> {
  const q = query(meetingsRef(uid), orderBy('createdAt', 'desc'));
  const snap = await getDocs(q);
  return snap.docs.map(d => d.data() as Meeting);
}

export async function saveMeetingCloud(uid: string, meeting: Meeting): Promise<void> {
  await setDoc(doc(meetingsRef(uid), meeting.id), meeting);
}

export async function deleteMeetingCloud(uid: string, id: string): Promise<void> {
  await deleteDoc(doc(meetingsRef(uid), id));
}

export async function updateMeetingSiteCloud(uid: string, id: string, siteName: string): Promise<void> {
  await updateDoc(doc(meetingsRef(uid), id), { siteName });
}

export async function renameSiteCloud(uid: string, oldName: string, newName: string, meetings: Meeting[]): Promise<void> {
  const batch = writeBatch(db);
  meetings
    .filter(m => m.siteName === oldName)
    .forEach(m => batch.update(doc(meetingsRef(uid), m.id), { siteName: newName }));
  await batch.commit();
}

// ── 공유 기능 ──────────────────────────────────────────

function inboxRef(toEmail: string) {
  return collection(db, 'inboxes', encodeURIComponent(toEmail), 'items');
}

export async function shareMeeting(fromEmail: string, toEmail: string, meeting: Meeting): Promise<void> {
  const shareId = `${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
  await setDoc(doc(inboxRef(toEmail.toLowerCase().trim()), shareId), {
    ...meeting,
    shareId,
    fromEmail,
    sharedAt: Date.now(),
  });
}

export async function getReceivedMeetings(toEmail: string): Promise<SharedMeeting[]> {
  const snap = await getDocs(query(inboxRef(toEmail.toLowerCase().trim()), orderBy('sharedAt', 'desc')));
  return snap.docs.map(d => d.data() as SharedMeeting);
}

export async function dismissSharedMeeting(toEmail: string, shareId: string): Promise<void> {
  await deleteDoc(doc(inboxRef(toEmail.toLowerCase().trim()), shareId));
}
