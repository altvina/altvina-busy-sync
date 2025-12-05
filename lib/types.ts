/**
 * Shared TypeScript interfaces and types for iCloud to Exchange calendar sync
 */

export interface NormalizedEvent {
  uid: string;
  title: string;
  description: string;
  start: Date;
  end: Date;
  location?: string;
  calendarName: string;
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
  showAs: "busy";
  sensitivity: "private";
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
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
}

