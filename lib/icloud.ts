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
 * Discover calendars for the iCloud account
 * Returns calendars matching the specified names
 */
export async function discoverCalendars(
  username: string,
  password: string,
  targetCalendarNames: string[]
): Promise<ICloudCalendar[]> {
  try {
    // Step 1: Get current-user-principal
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
    
    // Extract current-user-principal href
    let principalHrefMatch = principalXml.match(
      /<d:current-user-principal[^>]*>\s*<d:href[^>]*>([^<]+)<\/d:href>\s*<\/d:current-user-principal>/
    );
    
    if (!principalHrefMatch) {
      principalHrefMatch = principalXml.match(
        /<current-user-principal[^>]*>\s*<href[^>]*>([^<]+)<\/href>\s*<\/current-user-principal>/
      );
    }

    if (!principalHrefMatch) {
      const preview = principalXml.substring(0, 1000);
      console.error("Principal response XML:", preview);
      throw new Error(
        `Could not find current-user-principal in response. Response preview: ${preview}`
      );
    }

    const principalHref = principalHrefMatch[1];
    const principalUrl = principalHref.startsWith('http') 
      ? principalHref 
      : `${ICLOUD_BASE_URL}${principalHref}`;

    // Step 2: Query the principal for calendar-home-set
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
    
    // Extract calendar-home-set href
    let calendarHomeMatch = calendarHomeXml.match(
      /<c:calendar-home-set[^>]*>\s*<d:href[^>]*>([^<]+)<\/d:href>\s*<\/c:calendar-home-set>/
    );
    
    if (!calendarHomeMatch) {
      calendarHomeMatch = calendarHomeXml.match(
        /<calendar-home-set[^>]*>\s*<href[^>]*>([^<]+)<\/href>\s*<\/calendar-home-set>/
      );
    }

    if (!calendarHomeMatch) {
      const preview = calendarHomeXml.substring(0, 1000);
      console.error("Calendar home response XML:", preview);
      throw new Error(
        `Could not find calendar home set. Response preview: ${preview}`
      );
    }

    let calendarHomeUrl = calendarHomeMatch[1];
    // Make sure it's a full URL
    if (!calendarHomeUrl.startsWith('http')) {
      calendarHomeUrl = `${ICLOUD_BASE_URL}${calendarHomeUrl}`;
    }

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
    
    console.log(`Calendar discovery response received (${calendarsXml.length} chars)`);

    // Parse calendar list from XML response
    // Match calendar entries with display names
    const calendarMatches = calendarsXml.matchAll(
      /<d:response[^>]*>[\s\S]*?<d:href>([^<]+)<\/d:href>[\s\S]*?<d:displayname>([^<]+)<\/d:displayname>[\s\S]*?<\/d:response>/g
    );
    
    const allDiscoveredCalendars: Array<{name: string, url: string}> = [];

    for (const match of calendarMatches) {
      const url = match[1];
      const name = match[2];
      allDiscoveredCalendars.push({ name, url });
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

