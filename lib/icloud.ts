/**
 * iCloud CalDAV client for fetching calendar events
 */

import ICAL from "ical.js";
import type { NormalizedEvent, ICloudCalendar } from "./types.js";

const ICLOUD_BASE_URL = "https://caldav.icloud.com";

/**
 * Get Basic Auth header for iCloud CalDAV requests
 */
function getAuthHeader(username: string, password: string): string {
  const credentials = Buffer.from(`${username}:${password}`).toString("base64");
  return `Basic ${credentials}`;
}

/**
 * Internal: fetch calendar home URL (shared by discoverCalendars and discoverAllCalendars)
 */
async function getCalendarHomeUrl(
  username: string,
  password: string
): Promise<string> {
  const principalResponse = await fetch(`${ICLOUD_BASE_URL}/`, {
    method: "PROPFIND",
    headers: {
      Authorization: getAuthHeader(username, password),
      Depth: "0",
      "Content-Type": "application/xml",
    },
    body: `<?xml version="1.0" encoding="UTF-8"?>
<d:propfind xmlns:d="DAV:" xmlns:cs="http://calendarserver.org/ns/">
  <d:prop>
    <cs:getctag/>
    <d:current-user-principal/>
  </d:prop>
</d:propfind>`,
  });
  if (!principalResponse.ok) {
    throw new Error(
      `Failed to discover principal: ${principalResponse.status} ${principalResponse.statusText}`
    );
  }
  const principalXml = await principalResponse.text();
  let principalHrefMatch = principalXml.match(
    /<d:current-user-principal[^>]*>\s*<d:href[^>]*>([^<]+)<\/d:href>\s*<\/d:current-user-principal>/
  );
  if (!principalHrefMatch) {
    principalHrefMatch = principalXml.match(
      /<current-user-principal[^>]*>\s*<href[^>]*>([^<]+)<\/href>\s*<\/current-user-principal>/
    );
  }
  if (!principalHrefMatch) {
    throw new Error(
      `Could not find current-user-principal. Response preview: ${principalXml.substring(0, 500)}`
    );
  }
  const principalHref = principalHrefMatch[1];
  const principalUrl = principalHref.startsWith("http")
    ? principalHref
    : `${ICLOUD_BASE_URL}${principalHref}`;

  const calendarHomeResponse = await fetch(principalUrl, {
    method: "PROPFIND",
    headers: {
      Authorization: getAuthHeader(username, password),
      Depth: "0",
      "Content-Type": "application/xml",
    },
    body: `<?xml version="1.0" encoding="UTF-8"?>
<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">
  <d:prop>
    <c:calendar-home-set/>
  </d:prop>
</d:propfind>`,
  });
  if (!calendarHomeResponse.ok) {
    throw new Error(
      `Failed to get calendar home set: ${calendarHomeResponse.status} ${calendarHomeResponse.statusText}`
    );
  }
  const calendarHomeXml = await calendarHomeResponse.text();
  let calendarHomeMatch = calendarHomeXml.match(
    /<c:calendar-home-set[^>]*>\s*<d:href[^>]*>([^<]+)<\/d:href>\s*<\/c:calendar-home-set>/
  );
  if (!calendarHomeMatch) {
    calendarHomeMatch = calendarHomeXml.match(
      /<calendar-home-set[^>]*>\s*<href[^>]*>([^<]+)<\/href>\s*<\/calendar-home-set>/
    );
  }
  if (!calendarHomeMatch) {
    throw new Error(
      `Could not find calendar home set. Response preview: ${calendarHomeXml.substring(0, 500)}`
    );
  }
  let calendarHomeUrl = calendarHomeMatch[1];
  if (!calendarHomeUrl.startsWith("http")) {
    calendarHomeUrl = `${ICLOUD_BASE_URL}${calendarHomeUrl}`;
  }
  return calendarHomeUrl;
}

/**
 * Discover ALL calendars for the iCloud account (except Holidays, Birthdays, Reminders).
 * Use this to get exact names for ICLOUD_CAL1 / ICLOUD_CAL2 when write-back fails.
 */
export async function discoverAllCalendars(
  username: string,
  password: string
): Promise<ICloudCalendar[]> {
  const calendarHomeUrl = await getCalendarHomeUrl(username, password);
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
  const result: ICloudCalendar[] = [];
  const excludedNames = ["Holidays", "Birthdays", "Reminders"];
  const calendarMatches = calendarsXml.matchAll(
    /<response[^>]*>[\s\S]*?<href[^>]*>([^<]+)<\/href>[\s\S]*?<displayname[^>]*>([^<]+)<\/displayname>[\s\S]*?<\/response>/g
  );
  for (const match of calendarMatches) {
    let url = match[1];
    const name = match[2];
    if (!url.startsWith("http")) url = `${ICLOUD_BASE_URL}${url}`;
    const isPrincipal = url.endsWith("/calendars/") && !url.match(/\/calendars\/[^/]+\/$/);
    if (isPrincipal) continue;
    if (excludedNames.includes(name)) continue;
    result.push({ name, url });
  }
  return result;
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
    const calendarHomeUrl = await getCalendarHomeUrl(username, password);

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
    
    console.log(`Calendar discovery response received (${calendarsXml.length} chars)`);
    
    // Log a sample of the XML to debug parsing issues
    const xmlSample = calendarsXml.substring(0, 2000);
    console.log(`XML response sample (first 2000 chars):\n${xmlSample}`);

    // Parse calendar list from XML response
    // iCloud uses default namespace (xmlns="DAV:") so elements don't have prefixes
    // Structure: <response><href>...</href><propstat><prop><displayname>...</displayname></prop></propstat></response>
    let allDiscoveredCalendars: Array<{name: string, url: string}> = [];
    
    // Pattern for default namespace (xmlns="DAV:") - this is what iCloud uses
    // Match: <response>...<href>URL</href>...<displayname>NAME</displayname>...</response>
    const calendarMatches = calendarsXml.matchAll(
      /<response[^>]*>[\s\S]*?<href[^>]*>([^<]+)<\/href>[\s\S]*?<displayname[^>]*>([^<]+)<\/displayname>[\s\S]*?<\/response>/g
    );
    
    for (const match of calendarMatches) {
      let url = match[1];
      const name = match[2];
      
      // Convert relative URL to absolute URL if needed
      if (!url.startsWith('http')) {
        url = `${ICLOUD_BASE_URL}${url}`;
      }
      
      // Skip only the true principal entry (ends with exactly /calendars/ with nothing after)
      // Keep sub-paths like /calendars/home/ or /calendars/UUID/ as they are actual calendars
      // The principal entry is typically /272291719/calendars/ (user's calendar home)
      const isPrincipal = url.endsWith('/calendars/') && !url.match(/\/calendars\/[^\/]+\/$/);
      if (isPrincipal) {
        console.log(`Skipping principal/root entry: "${name}" (${url})`);
        continue;
      }
      
      allDiscoveredCalendars.push({ url, name });
    }

    
    console.log(`Total calendars discovered: ${allDiscoveredCalendars.length}`);
    console.log(`All discovered calendar names: ${allDiscoveredCalendars.map(c => `"${c.name}"`).join(", ")}`);
    console.log(`Looking for calendars matching: ${targetCalendarNames.map(n => `"${n}"`).join(", ")}`);

    for (const calendar of allDiscoveredCalendars) {
      const url = calendar.url;
      const name = calendar.name;

      // Filter out excluded calendars
      const excludedNames = ["Holidays", "Birthdays", "Reminders"];
      if (excludedNames.includes(name)) {
        console.log(`Excluding calendar: "${name}" (in excluded list)`);
        continue;
      }

      // Normalize names for comparison (trim and remove quotes)
      const normalizeName = (n: string) => n.trim().replace(/^["']|["']$/g, '');
      const normalizedTargetNames = targetCalendarNames.map(normalizeName);
      const normalizedName = normalizeName(name);

      // Only include calendars matching target names
      if (normalizedTargetNames.includes(normalizedName)) {
        console.log(`✓ Including calendar: "${name}" (matches target)`);
        calendars.push({ name, url });
      } else {
        console.log(`✗ Skipping calendar: "${name}" (does not match target names)`);
      }
    }
    
    console.log(`Matched ${calendars.length} calendar(s) out of ${allDiscoveredCalendars.length} total`);

    return calendars;
  } catch (error) {
    console.error("Error discovering calendars:", error);
    throw error;
  }
}

/**
 * Parse CalDAV REPORT multistatus into { href, icalText } per response (for eventUrl).
 */
function parseReportResponses(
  reportXml: string,
  calendarBaseUrl: string
): { href: string; icalText: string }[] {
  const blocks: { href: string; icalText: string }[] = [];
  const responseRegex = /<(?:d:)?response[^>]*>([\s\S]*?)<\/(?:d:)?response>/gi;
  let match: RegExpExecArray | null;
  while ((match = responseRegex.exec(reportXml)) !== null) {
    const block = match[1];
    const hrefMatch = block.match(/<d:href[^>]*>([^<]+)<\/d:href>/i) ?? block.match(/<href[^>]*>([^<]+)<\/href>/i);
    const dataMatch = block.match(/<c:calendar-data[^>]*>([\s\S]*?)<\/c:calendar-data>/i) ?? block.match(/<calendar-data[^>]*>([\s\S]*?)<\/calendar-data>/i);
    if (hrefMatch && dataMatch) {
      let href = hrefMatch[1].trim();
      if (!href.startsWith("http")) {
        href = href.startsWith("/") ? `${ICLOUD_BASE_URL}${href}` : `${calendarBaseUrl.replace(/\/?$/, "/")}${href}`;
      }
      blocks.push({ href, icalText: dataMatch[1] });
    }
  }
  return blocks;
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
    
    console.log(`Querying events from ${calendarName} from ${startStr} to ${endStr} (UTC)`);
    console.log(`Local time: ${start.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' })} to ${end.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' })} (PT)`);

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
    
    console.log(`CalDAV REPORT response for ${calendarName} (${reportXml.length} chars)`);
    if (reportXml.length < 500) {
      console.log(`Report XML: ${reportXml}`);
    } else {
      console.log(`Report XML sample (first 500 chars): ${reportXml.substring(0, 500)}`);
    }

    // Parse multistatus into response blocks (each has href + calendar-data) for eventUrl
    const responseBlocks = parseReportResponses(reportXml, calendarUrl);
    console.log(`Found ${responseBlocks.length} response block(s) in report`);

    const events: NormalizedEvent[] = [];

    for (const block of responseBlocks) {
      let icalText = block.icalText;
      const eventUrl = block.href;
      
      // Decode XML entities
      icalText = icalText
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&amp;/g, "&")
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'");
      
      // Remove CDATA markers if present
      icalText = icalText.replace(/<!\[CDATA\[/g, '').replace(/\]\]>/g, '');
      
      // Trim whitespace
      icalText = icalText.trim();
      
      // Skip empty blocks
      if (!icalText || icalText.length === 0) {
        console.log("Skipping empty calendar-data block");
        continue;
      }
      
      // Log first 200 chars for debugging
      console.log(`Parsing calendar-data block (${icalText.length} chars), sample: ${icalText.substring(0, 200)}`);

      try {
        // Check if this is an all-day event (VALUE=DATE format)
        const isAllDay = /DTSTART[^:]*VALUE=DATE/i.test(icalText) || /DTEND[^:]*VALUE=DATE/i.test(icalText);
        
        // Extract timezone and time info from raw iCalendar text
        let dtstartMatch = icalText.match(/DTSTART(?:;[^:]*TZID=([^:;]+))?:(\d{8}T\d{6})/i);
        let eventTimezone = dtstartMatch ? (dtstartMatch[1] || null) : null;
        let dtstartValue = dtstartMatch ? dtstartMatch[2] : null;
        
        if (isAllDay && !dtstartValue) {
          const dateOnlyMatch = icalText.match(/DTSTART[^:]*:(\d{8})/i);
          if (dateOnlyMatch) {
            dtstartValue = dateOnlyMatch[1];
            eventTimezone = null;
          }
        }
        
        let dtendMatch = icalText.match(/DTEND(?:;[^:]*TZID=([^:;]+))?:(\d{8}T\d{6})/i);
        let dtendValue = dtendMatch ? dtendMatch[2] : null;
        
        if (isAllDay && !dtendValue) {
          const dateOnlyMatch = icalText.match(/DTEND[^:]*:(\d{8})/i);
          if (dateOnlyMatch) {
            dtendValue = dateOnlyMatch[1];
          }
        }
        
        const jcalData = ICAL.parse(icalText);
        const comp = new ICAL.Component(jcalData);

        const vevents = comp.getAllSubcomponents("vevent");

        for (const vevent of vevents) {
          const normalized = normalizeEvent(vevent, calendarName, eventTimezone, dtstartValue, dtendValue, isAllDay, eventUrl);
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
  calendarName: string,
  eventTimezone: string | null,
  dtstartValue: string | null,
  dtendValue: string | null,
  isAllDay: boolean = false,
  eventUrl?: string
): NormalizedEvent | null {
  try {
    const event = new ICAL.Event(vevent);

    // Get UID from component (required for tracking)
    const uid = vevent.getFirstPropertyValue("uid") || "";
    
    // Get properties from Event object (more reliable)
    const summary = event.summary || "Untitled Event";
    const description = event.description || "";
    const location = event.location || undefined;

    // Get the ICAL.Time objects
    const startTime = event.startDate;
    const endTime = event.endDate;

    if (!startTime || !endTime) {
      console.warn("Event missing start or end date, skipping:", uid);
      return null;
    }

    // Convert to JavaScript Date objects
    let startDate = startTime.toJSDate();
    let endDate = endTime.toJSDate();

    // For all-day events, ensure they're at UTC midnight
    // All-day events in iCalendar are date-only and should remain as UTC midnight
    if (isAllDay) {
      // Ensure start is at UTC midnight of the start date
      startDate = new Date(Date.UTC(
        startDate.getUTCFullYear(),
        startDate.getUTCMonth(),
        startDate.getUTCDate(),
        0, 0, 0, 0
      ));
      // End date for all-day events should be UTC midnight of the day AFTER
      // (iCalendar all-day events end at start of next day)
      endDate = new Date(Date.UTC(
        endDate.getUTCFullYear(),
        endDate.getUTCMonth(),
        endDate.getUTCDate(),
        0, 0, 0, 0
      ));
      // If end date is same as start, it's a single-day event, so end should be next day
      if (endDate.getTime() === startDate.getTime()) {
        endDate = new Date(startDate.getTime() + 24 * 60 * 60 * 1000);
      }
      console.log(`All-day event "${summary}": ${startDate.toISOString()} to ${endDate.toISOString()}`);
    }
    
    // Fix timezone conversion issue with ical.js (only for timed events, not all-day)
    // Problem: toJSDate() sometimes treats timezone-aware times as UTC
    // Example: DTSTART;TZID=America/Los_Angeles:20251205T093000 (9:30 AM PT)
    // Should become: 2025-12-05T17:30:00.000Z (17:30 UTC, since PT is UTC-8)
    // But toJSDate() might return: 2025-12-05T09:30:00.000Z (treating 09:30 as UTC)
    if (!isAllDay && eventTimezone && eventTimezone !== 'UTC' && eventTimezone.includes('America/Los_Angeles') && dtstartValue && dtstartValue.includes('T')) {
      // Parse the raw iCalendar time value: 20251205T093000
      // Extract hour to see what the original time was
      const timePart = dtstartValue.split('T')[1]; // "093000"
      const originalHour = parseInt(timePart.substring(0, 2)); // 9
      const originalMinute = parseInt(timePart.substring(2, 4)); // 30
      
      // Get what UTC hour we got from toJSDate()
      const utcHour = startDate.getUTCHours();
      const utcMinute = startDate.getUTCMinutes();
      
      // If UTC hour matches original hour, toJSDate() didn't apply timezone
      // Original: 9:30 AM PT should become 17:30 UTC, not 09:30 UTC
      if (utcHour === originalHour && utcMinute === originalMinute) {
        // toJSDate() didn't apply timezone - add PST offset (UTC-8 means add 8 hours)
        const pstOffsetMs = 8 * 60 * 60 * 1000; // 8 hours in milliseconds
        startDate = new Date(startDate.getTime() + pstOffsetMs);
        console.log(`Fixed start timezone conversion for "${summary}": ${originalHour}:${String(originalMinute).padStart(2, '0')} PT -> added 8 hours PST offset`);
      }
      
      // Fix end time if we have the raw value
      if (dtendValue) {
        const endTimePart = dtendValue.split('T')[1];
        const originalEndHour = parseInt(endTimePart.substring(0, 2));
        const originalEndMinute = parseInt(endTimePart.substring(2, 4));
        const utcEndHour = endDate.getUTCHours();
        const utcEndMinute = endDate.getUTCMinutes();
        
        if (utcEndHour === originalEndHour && utcEndMinute === originalEndMinute) {
          const pstOffsetMs = 8 * 60 * 60 * 1000;
          endDate = new Date(endDate.getTime() + pstOffsetMs);
          console.log(`Fixed end timezone conversion for "${summary}": ${originalEndHour}:${String(originalEndMinute).padStart(2, '0')} PT -> added 8 hours PST offset`);
        }
      }
    }

    // Log for debugging
    console.log(`Event "${summary}": timezone=${eventTimezone || 'UTC'}, JSDate UTC=${startDate.toISOString()}, PT=${startDate.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' })}`);

    return {
      uid,
      title: summary,
      description,
      start: startDate,
      end: endDate,
      location,
      calendarName,
      isAllDay,
      eventUrl,
    };
  } catch (error) {
    console.error("Error normalizing event:", error);
    return null;
  }
}

/**
 * Build iCalendar (.ics) string from NormalizedEvent for CalDAV PUT
 */
function buildIcsFromEvent(event: NormalizedEvent, uidOverride?: string): string {
  const uid = uidOverride ?? event.uid;
  const formatUtc = (d: Date) => {
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, "0");
    const day = String(d.getUTCDate()).padStart(2, "0");
    const h = String(d.getUTCHours()).padStart(2, "0");
    const min = String(d.getUTCMinutes()).padStart(2, "0");
    const s = String(d.getUTCSeconds()).padStart(2, "0");
    return `${y}${m}${day}T${h}${min}${s}Z`;
  };
  const now = formatUtc(new Date());
  const escape = (s: string) => s.replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
  let dtStart: string;
  let dtEnd: string;
  if (event.isAllDay) {
    const d = (date: Date) => `${date.getUTCFullYear()}${String(date.getUTCMonth() + 1).padStart(2, "0")}${String(date.getUTCDate()).padStart(2, "0")}`;
    dtStart = `DTSTART;VALUE=DATE:${d(event.start)}`;
    dtEnd = `DTEND;VALUE=DATE:${d(event.end)}`;
  } else {
    dtStart = `DTSTART:${formatUtc(event.start)}`;
    dtEnd = `DTEND:${formatUtc(event.end)}`;
  }
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Altvina Busy Sync//EN",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${now}`,
    dtStart,
    dtEnd,
    `SUMMARY:${escape(event.title)}`,
    "TRANSP:OPAQUE",
  ];
  if (event.description) {
    lines.push(`DESCRIPTION:${escape(event.description)}`);
  }
  if (event.location) {
    lines.push(`LOCATION:${escape(event.location)}`);
  }
  lines.push("END:VEVENT", "END:VCALENDAR");
  return lines.join("\r\n");
}

/**
 * Create a new event on an iCloud CalDAV calendar (CAL1/CAL2 only).
 * Returns the UID used (generated or event.uid).
 */
export async function createIcloudEvent(
  username: string,
  password: string,
  calendarUrl: string,
  event: NormalizedEvent
): Promise<string> {
  const uid = event.uid || `${crypto.randomUUID().replace(/-/g, "")}@altvina-sync`;
  const eventToUse = { ...event, uid };
  const ical = buildIcsFromEvent(eventToUse);
  const eventFilename = `${uid.replace(/[^a-zA-Z0-9@._-]/g, "_")}.ics`;
  const url = calendarUrl.replace(/\/?$/, "/") + eventFilename;

  const response = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: getAuthHeader(username, password),
      "Content-Type": 'text/calendar; charset="utf-8"',
      "If-None-Match": "*",
    },
    body: ical,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Failed to create iCloud event: ${response.status} ${response.statusText} - ${text}`);
  }
  console.log(`Created iCloud event: ${event.title} (${uid})`);
  return uid;
}

/**
 * Update an existing event on iCloud CalDAV by its event URL.
 */
export async function updateIcloudEvent(
  username: string,
  password: string,
  eventUrl: string,
  event: NormalizedEvent
): Promise<void> {
  const ical = buildIcsFromEvent(event);

  const response = await fetch(eventUrl, {
    method: "PUT",
    headers: {
      Authorization: getAuthHeader(username, password),
      "Content-Type": 'text/calendar; charset="utf-8"',
    },
    body: ical,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Failed to update iCloud event: ${response.status} ${response.statusText} - ${text}`);
  }
  console.log(`Updated iCloud event: ${event.title} (${event.uid})`);
}

/**
 * Delete an event on iCloud CalDAV by its event URL.
 */
export async function deleteIcloudEventByUrl(
  username: string,
  password: string,
  eventUrl: string
): Promise<void> {
  const response = await fetch(eventUrl, {
    method: "DELETE",
    headers: {
      Authorization: getAuthHeader(username, password),
    },
  });

  if (!response.ok && response.status !== 404) {
    const text = await response.text();
    throw new Error(`Failed to delete iCloud event: ${response.status} ${response.statusText} - ${text}`);
  }
  console.log(`Deleted iCloud event at ${eventUrl}`);
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
  console.log(`Discovering calendars matching: ${calendarNames.join(", ")}`);
  const calendars = await discoverCalendars(username, password, calendarNames);

  if (calendars.length === 0) {
    console.warn(
      `No matching calendars found for: ${calendarNames.join(", ")}`
    );
    return [];
  }

  console.log(`Found ${calendars.length} matching calendar(s): ${calendars.map(c => c.name).join(", ")}`);
  console.log(`Fetching events from ${start.toISOString()} to ${end.toISOString()}`);

  const allEvents: NormalizedEvent[] = [];

  for (const calendar of calendars) {
    try {
      console.log(`Fetching events from calendar: ${calendar.name}`);
      const events = await fetchEvents(
        username,
        password,
        calendar.url,
        calendar.name,
        start,
        end
      );
      console.log(`Found ${events.length} event(s) in calendar: ${calendar.name}`);
      allEvents.push(...events);
    } catch (error) {
      console.error(
        `Failed to fetch events from calendar ${calendar.name}:`,
        error
      );
      // Continue with other calendars
    }
  }

  console.log(`Total events fetched: ${allEvents.length}`);
  return allEvents;
}

/**
 * Fetch events from a public calendar URL (webcal:// or https://)
 * Public calendars are typically .ics files that can be fetched via HTTP
 */
export async function fetchPublicCalendarEvents(
  calendarUrl: string,
  calendarName: string,
  start: Date,
  end: Date
): Promise<NormalizedEvent[]> {
  try {
    // Convert webcal:// to https://
    const httpUrl = calendarUrl.replace(/^webcal:\/\//i, "https://");
    
    console.log(`Fetching public calendar from: ${httpUrl}`);
    
    const response = await fetch(httpUrl);
    
    if (!response.ok) {
      throw new Error(
        `Failed to fetch public calendar: ${response.status} ${response.statusText}`
      );
    }
    
    const icalText = await response.text();
    console.log(`Fetched public calendar .ics file (${icalText.length} chars)`);
    
    // Parse the iCalendar data
    const jcalData = ICAL.parse(icalText);
    const comp = new ICAL.Component(jcalData);
    
    // Get all VEVENT components
    const vevents = comp.getAllSubcomponents("vevent");
    console.log(`Found ${vevents.length} events in public calendar`);
    
    const events: NormalizedEvent[] = [];
    
    for (const vevent of vevents) {
      try {
        // Check if this is an all-day event
        const rawText = vevent.toString();
        const isAllDay = /DTSTART[^:]*VALUE=DATE/i.test(rawText) || /DTEND[^:]*VALUE=DATE/i.test(rawText);
        
        // Extract timezone from DTSTART
        const dtstartMatch = rawText.match(/DTSTART(?:;[^:]*TZID=([^:;]+))?:(\d{8}T?\d{0,6})/i);
        const eventTimezone = dtstartMatch ? (dtstartMatch[1] || null) : null;
        const dtstartValue = dtstartMatch ? dtstartMatch[2] : null;
        
        // Extract DTEND
        const dtendMatch = rawText.match(/DTEND(?:;[^:]*TZID=([^:;]+))?:(\d{8}T?\d{0,6})/i);
        const dtendValue = dtendMatch ? dtendMatch[2] : null;
        
        const normalized = normalizeEvent(vevent, calendarName, eventTimezone, dtstartValue, dtendValue, isAllDay);
        
        if (normalized) {
          // Filter events to only include those within the sync window
          if (normalized.start <= end && normalized.end >= start) {
            events.push(normalized);
          }
        }
      } catch (error) {
        console.error("Error parsing event from public calendar:", error);
        // Continue with next event
      }
    }
    
    console.log(`Filtered to ${events.length} events within sync window`);
    return events;
  } catch (error) {
    console.error(`Error fetching public calendar from ${calendarUrl}:`, error);
    throw error;
  }
}

