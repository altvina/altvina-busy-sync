/**
 * iCloud CalDAV client for fetching calendar events
 */

import ICAL from "ical.js";
import type { NormalizedEvent, ICloudCalendar } from "./types";

const ICLOUD_BASE_URL = "https://caldav.icloud.com";

/**
 * Get Basic Auth header for iCloud CalDAV requests
 */
function getAuthHeader(username: string, password: string): string {
  const credentials = Buffer.from(`${username}:${password}`).toString("base64");
  return `Basic ${credentials}`;
}

/**
 * Discover calendars for the iCloud account
 * Returns calendars matching the specified names
 */
export async function discoverCalendars(
  username: string,
  password: string,
  targetCalendarNames: string[]
): Promise<ICloudCalendar[]> {
  try {
    // First, discover the principal URL
    const principalResponse = await fetch(`${ICLOUD_BASE_URL}/`, {
      method: "PROPFIND",
      headers: {
        Authorization: getAuthHeader(username, password),
        Depth: "0",
        "Content-Type": "application/xml",
      },
      body: `<?xml version="1.0" encoding="UTF-8"?>
<d:propfind xmlns:d="DAV:" xmlns:cd="urn:ietf:params:xml:ns:caldav">
  <d:prop>
    <cd:calendar-home-set/>
  </d:prop>
</d:propfind>`,
    });

    if (!principalResponse.ok) {
      throw new Error(
        `Failed to discover principal: ${principalResponse.status} ${principalResponse.statusText}`
      );
    }

    const principalXml = await principalResponse.text();
    const calendarHomeMatch = principalXml.match(
      /<cd:calendar-home-set[^>]*><d:href>([^<]+)<\/d:href><\/cd:calendar-home-set>/
    );

    if (!calendarHomeMatch) {
      throw new Error("Could not find calendar home set in principal response");
    }

    const calendarHomeUrl = calendarHomeMatch[1];

    // Now discover all calendars
    const calendarsResponse = await fetch(calendarHomeUrl, {
      method: "PROPFIND",
      headers: {
        Authorization: getAuthHeader(username, password),
        Depth: "1",
        "Content-Type": "application/xml",
      },
      body: `<?xml version="1.0" encoding="UTF-8"?>
<d:propfind xmlns:d="DAV:" xmlns:cd="urn:ietf:params:xml:ns:caldav" xmlns:c="urn:ietf:params:xml:ns:caldav">
  <d:prop>
    <d:displayname/>
    <cd:calendar-description/>
  </d:prop>
</d:propfind>`,
    });

    if (!calendarsResponse.ok) {
      throw new Error(
        `Failed to discover calendars: ${calendarsResponse.status} ${calendarsResponse.statusText}`
      );
    }

    const calendarsXml = await calendarsResponse.text();
    const calendars: ICloudCalendar[] = [];

    // Parse calendar list from XML response
    // Match calendar entries with display names
    const calendarMatches = calendarsXml.matchAll(
      /<d:response[^>]*>[\s\S]*?<d:href>([^<]+)<\/d:href>[\s\S]*?<d:displayname>([^<]+)<\/d:displayname>[\s\S]*?<\/d:response>/g
    );

    for (const match of calendarMatches) {
      const url = match[1];
      const name = match[2];

      // Filter out excluded calendars
      const excludedNames = ["Holidays", "Birthdays", "Reminders"];
      if (excludedNames.includes(name)) {
        continue;
      }

      // Only include calendars matching target names
      if (targetCalendarNames.includes(name)) {
        calendars.push({ name, url });
      }
    }

    return calendars;
  } catch (error) {
    console.error("Error discovering calendars:", error);
    throw error;
  }
}

/**
 * Fetch events from a specific iCloud calendar within a time window
 */
export async function fetchEvents(
  username: string,
  password: string,
  calendarUrl: string,
  calendarName: string,
  start: Date,
  end: Date
): Promise<NormalizedEvent[]> {
  try {
    // Format dates for CalDAV query (UTC)
    const formatDate = (date: Date): string => {
      return date.toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";
    };

    const startStr = formatDate(start);
    const endStr = formatDate(end);

    // CalDAV REPORT request for events in time range
    const reportResponse = await fetch(calendarUrl, {
      method: "REPORT",
      headers: {
        Authorization: getAuthHeader(username, password),
        Depth: "1",
        "Content-Type": "application/xml",
      },
      body: `<?xml version="1.0" encoding="UTF-8"?>
<c:calendar-query xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">
  <d:prop>
    <d:getetag/>
    <c:calendar-data/>
  </d:prop>
  <c:filter>
    <c:comp-filter name="VCALENDAR">
      <c:comp-filter name="VEVENT">
        <c:time-range start="${startStr}" end="${endStr}"/>
      </c:comp-filter>
    </c:comp-filter>
  </c:filter>
</c:calendar-query>`,
    });

    if (!reportResponse.ok) {
      throw new Error(
        `Failed to fetch events: ${reportResponse.status} ${reportResponse.statusText}`
      );
    }

    const reportXml = await reportResponse.text();

    // Extract calendar-data from XML response
    const calendarDataMatches = reportXml.matchAll(
      /<c:calendar-data[^>]*>([\s\S]*?)<\/c:calendar-data>/g
    );

    const events: NormalizedEvent[] = [];

    for (const match of calendarDataMatches) {
      const icalText = match[1]
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&amp;/g, "&");

      try {
        const jcalData = ICAL.parse(icalText);
        const comp = new ICAL.Component(jcalData);

        // Get all VEVENT components
        const vevents = comp.getAllSubcomponents("vevent");

        for (const vevent of vevents) {
          const normalized = normalizeEvent(vevent, calendarName);
          if (normalized) {
            events.push(normalized);
          }
        }
      } catch (parseError) {
        console.error("Error parsing iCalendar data:", parseError);
        // Continue with next event
      }
    }

    return events;
  } catch (error) {
    console.error(`Error fetching events from ${calendarName}:`, error);
    throw error;
  }
}

/**
 * Normalize a VEVENT component to our standard format
 */
function normalizeEvent(
  vevent: ICAL.Component,
  calendarName: string
): NormalizedEvent | null {
  try {
    const event = new ICAL.Event(vevent);

    // Get UID from component (required for tracking)
    const uid = vevent.getFirstPropertyValue("uid") || "";
    
    // Get properties from Event object (more reliable)
    const summary = event.summary || "Untitled Event";
    const description = event.description || "";
    const location = event.location || undefined;

    const startDate = event.startDate?.toJSDate();
    const endDate = event.endDate?.toJSDate();

    if (!startDate || !endDate) {
      console.warn("Event missing start or end date, skipping:", uid);
      return null;
    }

    return {
      uid,
      title: summary,
      description,
      start: startDate,
      end: endDate,
      location,
      calendarName,
    };
  } catch (error) {
    console.error("Error normalizing event:", error);
    return null;
  }
}

/**
 * Fetch all events from specified iCloud calendars
 */
export async function fetchAllEvents(
  username: string,
  password: string,
  calendarNames: string[],
  start: Date,
  end: Date
): Promise<NormalizedEvent[]> {
  const calendars = await discoverCalendars(username, password, calendarNames);

  if (calendars.length === 0) {
    console.warn(
      `No matching calendars found for: ${calendarNames.join(", ")}`
    );
    return [];
  }

  const allEvents: NormalizedEvent[] = [];

  for (const calendar of calendars) {
    try {
      const events = await fetchEvents(
        username,
        password,
        calendar.url,
        calendar.name,
        start,
        end
      );
      allEvents.push(...events);
    } catch (error) {
      console.error(
        `Failed to fetch events from calendar ${calendar.name}:`,
        error
      );
      // Continue with other calendars
    }
  }

  return allEvents;
}

