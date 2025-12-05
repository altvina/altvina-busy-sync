/**
 * Microsoft Graph API client for Exchange calendar operations
 */

import type {
  GraphEvent,
  GraphCalendar,
  GraphEventResponse,
  NormalizedEvent,
} from "./types.js";

export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

/**
 * Get OAuth access token using client credentials flow
 */
export async function getAccessToken(
  tenantId: string,
  clientId: string,
  clientSecret: string
): Promise<string> {
  try {
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    const response = await fetch(tokenUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: params.toString(),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to get access token: ${response.status} ${response.statusText} - ${errorText}`
      );
    }

    const data = (await response.json()) as { access_token: string; error?: any };
    
    if (data.error) {
      console.error("Token request error:", data.error);
      throw new Error(
        `Failed to get access token: ${data.error.error_description || data.error.message || JSON.stringify(data.error)}`
      );
    }
    
    if (!data.access_token) {
      throw new Error("Access token not found in response");
    }
    
    return data.access_token;
  } catch (error) {
    console.error("Error getting access token:", error);
    throw error;
  }
}

/**
 * Find calendar by name for a specific user
 */
export async function findTargetCalendar(
  accessToken: string,
  userId: string,
  calendarName: string
): Promise<GraphCalendar> {
  try {
    const url = `${GRAPH_BASE_URL}/users/${userId}/calendars`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`Graph API error details:`, {
        status: response.status,
        statusText: response.statusText,
        url,
        userId,
        errorText,
      });
      
      // Provide helpful error message for common issues
      if (response.status === 403) {
        throw new Error(
          `Access denied to calendars for user ${userId}. ` +
          `Please verify: 1) Calendars.ReadWrite is set as APPLICATION permission (not Delegated), ` +
          `2) Admin consent has been granted, 3) The user ID is correct. ` +
          `Error: ${errorText}`
        );
      }
      
      throw new Error(
        `Failed to list calendars: ${response.status} ${response.statusText} - ${errorText}`
      );
    }

    const data = (await response.json()) as { value: GraphCalendar[] };
    const calendars: GraphCalendar[] = data.value || [];

    // Trim and normalize calendar names for comparison
    const normalizeName = (name: string): string => {
      return name.trim().replace(/^["']|["']$/g, '');
    };

    const normalizedTargetName = normalizeName(calendarName);
    
    console.log(`Looking for calendar: "${normalizedTargetName}"`);
    console.log(`Available calendars: ${calendars.map((c) => `"${c.name}"`).join(", ")}`);

    const targetCalendar = calendars.find(
      (cal) => normalizeName(cal.name) === normalizedTargetName
    );

    if (!targetCalendar) {
      throw new Error(
        `Calendar "${normalizedTargetName}" not found. Available calendars: ${calendars.map((c) => c.name).join(", ")}`
      );
    }
    
    console.log(`Found target calendar: "${targetCalendar.name}" (ID: ${targetCalendar.id})`);

    return targetCalendar;
  } catch (error) {
    console.error("Error finding target calendar:", error);
    throw error;
  }
}

/**
 * List all events in a calendar within a time window (with pagination support)
 */
export async function listEventsInWindow(
  accessToken: string,
  userId: string,
  calendarId: string,
  start: Date,
  end: Date
): Promise<GraphEventResponse[]> {
  const formatDate = (date: Date): string => {
    return date.toISOString();
  };

  const startStr = formatDate(start);
  const endStr = formatDate(end);
  const allEvents: GraphEventResponse[] = [];
  let nextLink: string | null = null;

  do {
    let url: string;
    if (nextLink) {
      url = nextLink;
    } else {
      // Request iCalUId and showAs fields explicitly to match events by UID and preserve status
      url = `${GRAPH_BASE_URL}/users/${userId}/calendars/${calendarId}/calendarView?startDateTime=${encodeURIComponent(startStr)}&endDateTime=${encodeURIComponent(endStr)}&$top=100&$select=id,subject,iCalUId,showAs,start,end`;
    }

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to query events: ${response.status} ${response.statusText} - ${errorText}`
      );
    }

    const data = (await response.json()) as { 
      value: GraphEventResponse[];
      "@odata.nextLink"?: string;
    };
    
    allEvents.push(...(data.value || []));
    nextLink = data["@odata.nextLink"] || null;
  } while (nextLink);

  return allEvents;
}

/**
 * Delete all events in a calendar within the specified time window
 * Handles pagination to ensure all events are deleted
 */
export async function deleteEventsInWindow(
  accessToken: string,
  userId: string,
  calendarId: string,
  start: Date,
  end: Date,
  timezone: string
): Promise<number> {
  try {
    const startStr = start.toISOString();
    const endStr = end.toISOString();

    console.log(`Querying events to delete in window ${startStr} to ${endStr}`);
    
    // Get all events (with pagination)
    const events = await listEventsInWindow(accessToken, userId, calendarId, start, end);

    console.log(`Found ${events.length} event(s) to delete`);
    
    // Log event details for debugging duplicates
    if (events.length > 0) {
      const eventSummary = events.map(e => `"${e.subject}" (${e.start.dateTime})`).slice(0, 10);
      console.log(`Events to delete (showing first 10): ${eventSummary.join(", ")}${events.length > 10 ? ` ... and ${events.length - 10} more` : ""}`);
      
      // Check for duplicates by subject and time
      const eventMap = new Map<string, number>();
      events.forEach(e => {
        const key = `${e.subject}|${e.start.dateTime}`;
        eventMap.set(key, (eventMap.get(key) || 0) + 1);
      });
      const duplicates = Array.from(eventMap.entries()).filter(([_, count]) => count > 1);
      if (duplicates.length > 0) {
        console.warn(`Found ${duplicates.length} duplicate event pattern(s) in calendar: ${duplicates.map(([key]) => key).join(", ")}`);
      }
    }

    // Delete each event
    let deletedCount = 0;
    let failedDeletes: string[] = [];
    
    for (const event of events) {
      try {
        const deleteUrl = `${GRAPH_BASE_URL}/users/${userId}/calendars/${calendarId}/events/${event.id}`;
        const deleteResponse = await fetch(deleteUrl, {
          method: "DELETE",
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        });

        if (deleteResponse.ok) {
          deletedCount++;
        } else {
          const errorText = await deleteResponse.text();
          console.warn(
            `Failed to delete event "${event.subject}" (${event.id}): ${deleteResponse.status} - ${errorText}`
          );
          failedDeletes.push(event.id);
        }
      } catch (error) {
        console.error(`Error deleting event "${event.subject}" (${event.id}):`, error);
        failedDeletes.push(event.id);
      }
    }
    
    if (failedDeletes.length > 0) {
      console.warn(`Failed to delete ${failedDeletes.length} event(s): ${failedDeletes.join(", ")}`);
    }

    console.log(`Successfully deleted ${deletedCount} of ${events.length} event(s)`);
    return deletedCount;
  } catch (error) {
    console.error("Error deleting events in window:", error);
    throw error;
  }
}

/**
 * Create a new event in the Exchange calendar
 */
export async function createEvent(
  accessToken: string,
  userId: string,
  calendarId: string,
  event: NormalizedEvent,
  timezone: string
): Promise<void> {
  try {
    // Format date for Graph API in the specified timezone
    // Microsoft Graph expects dateTime in ISO 8601 format representing local time in the timezone
    const formatDateTime = (date: Date, tz: string): string => {
      // Create a formatter for the target timezone
      const formatter = new Intl.DateTimeFormat('en-CA', {
        timeZone: tz,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false,
      });
      
      // Format the date - en-CA gives us YYYY-MM-DD format
      const parts = formatter.formatToParts(date);
      const year = parts.find(p => p.type === 'year')!.value;
      const month = parts.find(p => p.type === 'month')!.value;
      const day = parts.find(p => p.type === 'day')!.value;
      const hour = parts.find(p => p.type === 'hour')!.value;
      const minute = parts.find(p => p.type === 'minute')!.value;
      const second = parts.find(p => p.type === 'second')!.value;
      
      return `${year}-${month}-${day}T${hour}:${minute}:${second}`;
    };
    
    // Log the conversion for debugging
    const startFormatted = formatDateTime(event.start, timezone);
    const endFormatted = formatDateTime(event.end, timezone);
    console.log(`Creating event "${event.title}": start UTC=${event.start.toISOString()} -> ${startFormatted} (${timezone})`);

    const graphEvent: GraphEvent = {
      subject: event.title,
      body: {
        contentType: "text",
        content: `${event.description}\n\nSynced UID: ${event.uid}`,
      },
      start: {
        dateTime: formatDateTime(event.start, timezone),
        timeZone: timezone,
      },
      end: {
        dateTime: formatDateTime(event.end, timezone),
        timeZone: timezone,
      },
      showAs: "busy",
      sensitivity: "private",
    };

    if (event.location) {
      graphEvent.location = {
        displayName: event.location,
      };
    }

    const url = `${GRAPH_BASE_URL}/users/${userId}/calendars/${calendarId}/events`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(graphEvent),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to create event: ${response.status} ${response.statusText} - ${errorText}`
      );
    }
  } catch (error) {
    console.error("Error creating event:", error);
    throw error;
  }
}

/**
 * Find an event by UID in the calendar
 */
async function findEventByUid(
  accessToken: string,
  userId: string,
  calendarId: string,
  uid: string,
  window: { start: Date; end: Date }
): Promise<GraphEventResponse | null> {
  try {
    const events = await listEventsInWindow(accessToken, userId, calendarId, window.start, window.end);
    // Find event with matching iCalUId (Microsoft Graph stores UID in iCalUId field)
    return events.find(e => e.iCalUId === uid) || null;
  } catch (error) {
    console.warn(`Error finding event with UID ${uid}:`, error);
    return null;
  }
}

/**
 * Update an existing event in the Exchange calendar
 * Preserves showAs status if it's "free" (user manually changed it)
 */
export async function updateEvent(
  accessToken: string,
  userId: string,
  calendarId: string,
  eventId: string,
  event: NormalizedEvent,
  timezone: string,
  preserveShowAs?: "free" | "tentative" | "busy" | "oof" | "workingElsewhere" | "unknown"
): Promise<void> {
  try {
    // Format date for Graph API in the specified timezone
    const formatDateTime = (date: Date, tz: string): string => {
      const formatter = new Intl.DateTimeFormat('en-CA', {
        timeZone: tz,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false,
      });
      
      const parts = formatter.formatToParts(date);
      const year = parts.find(p => p.type === 'year')!.value;
      const month = parts.find(p => p.type === 'month')!.value;
      const day = parts.find(p => p.type === 'day')!.value;
      const hour = parts.find(p => p.type === 'hour')!.value;
      const minute = parts.find(p => p.type === 'minute')!.value;
      const second = parts.find(p => p.type === 'second')!.value;
      
      return `${year}-${month}-${day}T${hour}:${minute}:${second}`;
    };

    const graphEvent: Partial<GraphEvent> = {
      subject: event.title,
      body: {
        contentType: "text",
        content: `${event.description}\n\nSynced UID: ${event.uid}`,
      },
      start: {
        dateTime: formatDateTime(event.start, timezone),
        timeZone: timezone,
      },
      end: {
        dateTime: formatDateTime(event.end, timezone),
        timeZone: timezone,
      },
      // Preserve showAs if user manually set it to "free", otherwise set to "busy"
      showAs: preserveShowAs === "free" ? "free" : "busy",
      sensitivity: "private",
    };

    if (event.location) {
      graphEvent.location = {
        displayName: event.location,
      };
    }

    const url = `${GRAPH_BASE_URL}/users/${userId}/calendars/${calendarId}/events/${eventId}`;

    const response = await fetch(url, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(graphEvent),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to update event: ${response.status} ${response.statusText} - ${errorText}`
      );
    }
    
    if (preserveShowAs === "free") {
      console.log(`Updated event "${event.title}" (preserved showAs: free)`);
    }
  } catch (error) {
    console.error("Error updating event:", error);
    throw error;
  }
}

/**
 * Sync events using update-or-create pattern
 * Preserves manual showAs="free" status changes
 * Returns counts of created, updated, and skipped events
 */
export async function syncEvents(
  accessToken: string,
  userId: string,
  calendarId: string,
  events: NormalizedEvent[],
  timezone: string,
  window: { start: Date; end: Date }
): Promise<{ created: number; updated: number; skipped: number }> {
  let createdCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  const processedUids = new Set<string>(); // Track UIDs in the batch

  // Get all existing events in the calendar to check for manual status changes
  const existingEvents = await listEventsInWindow(accessToken, userId, calendarId, window.start, window.end);
  const existingEventsByUid = new Map<string, GraphEventResponse>();
  existingEvents.forEach(e => {
    if (e.iCalUId) {
      existingEventsByUid.set(e.iCalUId, e);
    }
  });

  for (const event of events) {
    // Check for duplicate UIDs in the batch
    if (processedUids.has(event.uid)) {
      console.warn(`Skipping duplicate UID in batch: "${event.title}" (${event.uid})`);
      skippedCount++;
      continue;
    }
    processedUids.add(event.uid);
    
    // Check if event already exists in calendar
    const existingEvent = existingEventsByUid.get(event.uid);
    
    if (existingEvent) {
      // Event exists - update it, preserving showAs if it's "free"
      try {
        const preserveShowAs = existingEvent.showAs === "free" ? "free" : undefined;
        await updateEvent(
          accessToken,
          userId,
          calendarId,
          existingEvent.id,
          event,
          timezone,
          preserveShowAs
        );
        updatedCount++;
      } catch (error) {
        console.error(
          `Failed to update event "${event.title}" (${event.uid}):`,
          error
        );
        // Continue with next event
      }
    } else {
      // Event doesn't exist - create it
      try {
        await createEvent(accessToken, userId, calendarId, event, timezone);
        createdCount++;
      } catch (error) {
        console.error(
          `Failed to create event "${event.title}" (${event.uid}):`,
          error
        );
        // Continue with next event
      }
    }
  }

  return { created: createdCount, updated: updatedCount, skipped: skippedCount };
}


