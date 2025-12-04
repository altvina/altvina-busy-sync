/**
 * Microsoft Graph API client for Exchange calendar operations
 */

import type {
  GraphEvent,
  GraphCalendar,
  GraphEventResponse,
  NormalizedEvent,
} from "./types.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

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

    const data = (await response.json()) as { access_token: string };
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
      throw new Error(
        `Failed to list calendars: ${response.status} ${response.statusText} - ${errorText}`
      );
    }

    const data = (await response.json()) as { value: GraphCalendar[] };
    const calendars: GraphCalendar[] = data.value || [];

    const targetCalendar = calendars.find(
      (cal) => cal.name === calendarName
    );

    if (!targetCalendar) {
      throw new Error(
        `Calendar "${calendarName}" not found. Available calendars: ${calendars.map((c) => c.name).join(", ")}`
      );
    }

    return targetCalendar;
  } catch (error) {
    console.error("Error finding target calendar:", error);
    throw error;
  }
}

/**
 * Delete all events in a calendar within the specified time window
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
    // Format dates for Graph API (ISO 8601 UTC)
    const formatDate = (date: Date): string => {
      return date.toISOString();
    };

    const startStr = formatDate(start);
    const endStr = formatDate(end);

    // Use calendarView endpoint for reliable date range queries (handles recurring events)
    const url = `${GRAPH_BASE_URL}/users/${userId}/calendars/${calendarId}/calendarView?startDateTime=${encodeURIComponent(startStr)}&endDateTime=${encodeURIComponent(endStr)}`;

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

    const data = (await response.json()) as { value: GraphEventResponse[] };
    const events: GraphEventResponse[] = data.value || [];

    // Delete each event
    let deletedCount = 0;
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
          console.warn(
            `Failed to delete event ${event.id}: ${deleteResponse.status}`
          );
        }
      } catch (error) {
        console.error(`Error deleting event ${event.id}:`, error);
      }
    }

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
    // Format date for Graph API (ISO string with timezone)
    const formatDateTime = (date: Date): string => {
      return date.toISOString();
    };

    const graphEvent: GraphEvent = {
      subject: event.title,
      body: {
        contentType: "text",
        content: `${event.description}\n\nSynced UID: ${event.uid}`,
      },
      start: {
        dateTime: formatDateTime(event.start),
        timeZone: timezone,
      },
      end: {
        dateTime: formatDateTime(event.end),
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
 * Create multiple events in the Exchange calendar
 */
export async function createEvents(
  accessToken: string,
  userId: string,
  calendarId: string,
  events: NormalizedEvent[],
  timezone: string
): Promise<number> {
  let createdCount = 0;

  for (const event of events) {
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

  return createdCount;
}

