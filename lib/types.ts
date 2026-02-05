/**
 * Shared TypeScript interfaces and types for iCloud to Exchange calendar sync
 */

/** Sync source tag stored in Outlook body for two-way sync (CAL3 = read-only public). */
export type SyncSourceTag = "CAL1" | "CAL2" | "CAL3";

export interface NormalizedEvent {
  uid: string;
  title: string;
  description: string;
  start: Date;
  end: Date;
  location?: string;
  calendarName: string;
  isAllDay?: boolean; // True if event is all-day (VALUE=DATE in iCalendar)
  /** CalDAV event URL (when from CalDAV fetch) for update/delete write-back */
  eventUrl?: string;
}

export interface SyncWindow {
  start: Date;
  end: Date;
}

export interface ICloudCalendar {
  name: string;
  url: string;
}

export interface GraphEvent {
  subject: string;
  body: {
    contentType: "text";
    content: string;
  };
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  isAllDay?: boolean; // True for all-day events (use UTC midnight dates)
  showAs: "free" | "tentative" | "busy" | "oof" | "workingElsewhere" | "unknown";
  sensitivity: "private";
  iCalUId?: string; // iCalendar UID for matching events
  location?: {
    displayName: string;
  };
}

export interface GraphCalendar {
  id: string;
  name: string;
}

export interface GraphEventResponse {
  id: string;
  subject: string;
  iCalUId?: string; // iCalendar UID, used to match events
  showAs?: "free" | "tentative" | "busy" | "oof" | "workingElsewhere" | "unknown"; // Preserve manual status changes
  isAllDay?: boolean;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  /** Present when listEventsInWindow called with includeBody: true */
  body?: {
    contentType?: string;
    content?: string;
  };
}


