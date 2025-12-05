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

    // Extract calendar-data from XML response
    // iCloud uses default namespace, so try multiple patterns
    let matchesArray: RegExpMatchArray[] = [];
    
    // Try pattern 1: with c: namespace prefix
    let calendarDataMatches = reportXml.matchAll(
      /<c:calendar-data[^>]*>([\s\S]*?)<\/c:calendar-data>/g
    );
    matchesArray = Array.from(calendarDataMatches);
    
    // If no matches, try pattern 2: without namespace prefix (default namespace)
    if (matchesArray.length === 0) {
      console.log("Trying calendar-data pattern without namespace prefix...");
      calendarDataMatches = reportXml.matchAll(
        /<calendar-data[^>]*>([\s\S]*?)<\/calendar-data>/g
      );
      matchesArray = Array.from(calendarDataMatches);
    }
    
    // Try pattern 3: with xmlns attribute (CDATA might be escaped)
    if (matchesArray.length === 0) {
      console.log("Trying calendar-data pattern with xmlns...");
      calendarDataMatches = reportXml.matchAll(
        /<calendar-data[^>]*xmlns[^>]*>([\s\S]*?)<\/calendar-data>/g
      );
      matchesArray = Array.from(calendarDataMatches);
    }
    
    // Try pattern 4: flexible namespace with any prefix
    if (matchesArray.length === 0) {
      console.log("Trying calendar-data pattern with flexible namespace...");
      calendarDataMatches = reportXml.matchAll(
        /<[^:>]*:calendar-data[^>]*>([\s\S]*?)<\/[^:>]*:calendar-data>/g
      );
      matchesArray = Array.from(calendarDataMatches);
    }
    
    console.log(`Found ${matchesArray.length} calendar-data block(s) in report response`);
    
    // If still no matches, log a larger sample to see the actual structure
    if (matchesArray.length === 0 && reportXml.length > 1000) {
      console.log(`No calendar-data found. XML sample (chars 500-1500): ${reportXml.substring(500, 1500)}`);
    }

    const events: NormalizedEvent[] = [];

    for (const match of matchesArray) {
      let icalText = match[1];
      
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
        // Extract timezone and time info from raw iCalendar text
        // Look for DTSTART with timezone: DTSTART;TZID=America/Los_Angeles:20251205T093000
        const dtstartMatch = icalText.match(/DTSTART(?:;[^:]*TZID=([^:;]+))?:(\d{8}T\d{6})/i);
        const eventTimezone = dtstartMatch ? (dtstartMatch[1] || null) : null;
        const dtstartValue = dtstartMatch ? dtstartMatch[2] : null; // Format: 20251205T093000
        
        // Also extract DTEND for end time
        const dtendMatch = icalText.match(/DTEND(?:;[^:]*TZID=([^:;]+))?:(\d{8}T\d{6})/i);
        const dtendValue = dtendMatch ? dtendMatch[2] : null;
        
        const jcalData = ICAL.parse(icalText);
        const comp = new ICAL.Component(jcalData);

        // Get all VEVENT components
        const vevents = comp.getAllSubcomponents("vevent");

        for (const vevent of vevents) {
          const normalized = normalizeEvent(vevent, calendarName, eventTimezone, dtstartValue, dtendValue);
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
  dtendValue: string | null
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

    // Fix timezone conversion issue with ical.js
    // Problem: toJSDate() sometimes treats timezone-aware times as UTC
    // Example: DTSTART;TZID=America/Los_Angeles:20251205T093000 (9:30 AM PT)
    // Should become: 2025-12-05T17:30:00.000Z (17:30 UTC, since PT is UTC-8)
    // But toJSDate() might return: 2025-12-05T09:30:00.000Z (treating 09:30 as UTC)
    if (eventTimezone && eventTimezone !== 'UTC' && eventTimezone.includes('America/Los_Angeles') && dtstartValue) {
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


