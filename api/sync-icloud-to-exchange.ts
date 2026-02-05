/**
 * Vercel serverless function to sync iCloud calendars to Exchange calendar
 * Runs every 15 minutes via Vercel Cron Job
 */

import type { VercelRequest, VercelResponse } from "@vercel/node";
import { fetchAllEvents } from "../lib/icloud.js";
import {
  getAccessToken,
  findTargetCalendar,
  syncEvents,
  listEventsInWindow,
  GRAPH_BASE_URL,
} from "../lib/graph.js";
import type { SyncWindow } from "../lib/types.js";

/**
 * Calculate sync window based on lookback and lookahead days
 */
function calculateSyncWindow(
  lookbackDays: number,
  lookaheadDays: number
): SyncWindow {
  const now = new Date();
  const start = new Date(now);
  start.setDate(start.getDate() - lookbackDays);
  start.setHours(0, 0, 0, 0);

  const end = new Date(now);
  end.setDate(end.getDate() + lookaheadDays);
  end.setHours(23, 59, 59, 999);

  return { start, end };
}


/**
 * Validate required environment variables
 */
function validateEnvVars(): {
  icloudUsername: string;
  icloudPassword: string;
  icloudCal1: string;
  icloudCal2: string;
  icloudCal3?: string; // Optional public calendar URL
  msTenantId: string;
  msClientId: string;
  msClientSecret: string;
  msUserId: string;
  msTargetCalendarName: string;
  syncLookbackDays: number;
  syncLookaheadDays: number;
  timezone: string;
} {
  const requiredVars = {
    ICLOUD_USERNAME: process.env.ICLOUD_USERNAME,
    ICLOUD_APP_PASSWORD: process.env.ICLOUD_APP_PASSWORD,
    ICLOUD_CAL1: process.env.ICLOUD_CAL1,
    ICLOUD_CAL2: process.env.ICLOUD_CAL2,
    ICLOUD_CAL3: process.env.ICLOUD_CAL3, // Optional
    MS_TENANT_ID: process.env.MS_TENANT_ID,
    MS_CLIENT_ID: process.env.MS_CLIENT_ID,
    MS_CLIENT_SECRET: process.env.MS_CLIENT_SECRET,
    MS_USER_ID: process.env.MS_USER_ID,
    MS_TARGET_CALENDAR_NAME: process.env.MS_TARGET_CALENDAR_NAME,
    SYNC_LOOKBACK_DAYS: process.env.SYNC_LOOKBACK_DAYS,
    SYNC_LOOKAHEAD_DAYS: process.env.SYNC_LOOKAHEAD_DAYS,
    TIMEZONE: process.env.TIMEZONE,
  };

  // ICLOUD_CAL3 is optional
  const missing = Object.entries(requiredVars)
    .filter(([key, value]) => !value && key !== 'ICLOUD_CAL3')
    .map(([key]) => key);

  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(", ")}`);
  }

  const lookbackDays = parseInt(requiredVars.SYNC_LOOKBACK_DAYS!, 10);
  const lookaheadDays = parseInt(requiredVars.SYNC_LOOKAHEAD_DAYS!, 10);

  if (isNaN(lookbackDays) || isNaN(lookaheadDays)) {
    throw new Error(
      "SYNC_LOOKBACK_DAYS and SYNC_LOOKAHEAD_DAYS must be valid numbers"
    );
  }

  // Helper to trim quotes and whitespace from env vars
  const trimEnvVar = (value: string): string => {
    return value.trim().replace(/^["']|["']$/g, '');
  };

  return {
    icloudUsername: trimEnvVar(requiredVars.ICLOUD_USERNAME!),
    icloudPassword: trimEnvVar(requiredVars.ICLOUD_APP_PASSWORD!),
    icloudCal1: trimEnvVar(requiredVars.ICLOUD_CAL1!),
    icloudCal2: trimEnvVar(requiredVars.ICLOUD_CAL2!),
    icloudCal3: requiredVars.ICLOUD_CAL3 ? trimEnvVar(requiredVars.ICLOUD_CAL3) : undefined,
    msTenantId: trimEnvVar(requiredVars.MS_TENANT_ID!),
    msClientId: trimEnvVar(requiredVars.MS_CLIENT_ID!),
    msClientSecret: trimEnvVar(requiredVars.MS_CLIENT_SECRET!),
    msUserId: trimEnvVar(requiredVars.MS_USER_ID!),
    msTargetCalendarName: trimEnvVar(requiredVars.MS_TARGET_CALENDAR_NAME!),
    syncLookbackDays: lookbackDays,
    syncLookaheadDays: lookaheadDays,
    timezone: trimEnvVar(requiredVars.TIMEZONE!),
  };
}

/**
 * Main sync handler
 */
export default async function handler(
  req: VercelRequest,
  res: VercelResponse
) {
  const startTime = Date.now();

  try {
    console.log("Starting iCloud to Exchange calendar sync...");

    // Validate environment variables
    const env = validateEnvVars();

    // Calculate sync window
    const window = calculateSyncWindow(
      env.syncLookbackDays,
      env.syncLookaheadDays
    );
    console.log(
      `Sync window: ${window.start.toISOString()} to ${window.end.toISOString()}`
    );

    // Step 1: Fetch events from iCloud calendars
    const calendarList = [env.icloudCal1, env.icloudCal2];
    const calendarListStr = calendarList.join(", ");
    
    console.log(
      `Fetching events from CalDAV calendars: ${calendarListStr}`
    );
    const iCloudEvents = await fetchAllEvents(
      env.icloudUsername,
      env.icloudPassword,
      calendarList,
      window.start,
      window.end
    );
    console.log(`Fetched ${iCloudEvents.length} events from CalDAV calendars`);
    
    // Step 1b: Fetch events from public calendar URL if configured
    if (env.icloudCal3) {
      console.log(`Fetching events from public calendar URL: ${env.icloudCal3}`);
      const { fetchPublicCalendarEvents } = await import("../lib/icloud.js");
      const publicEvents = await fetchPublicCalendarEvents(
        env.icloudCal3,
        "Personal Calendar",
        window.start,
        window.end
      );
      console.log(`Fetched ${publicEvents.length} events from public calendar`);
      iCloudEvents.push(...publicEvents);
    }
    
    console.log(`Total events fetched from all sources: ${iCloudEvents.length}`);

    // Step 2: Get Microsoft Graph access token
    console.log("Getting Microsoft Graph access token...");
    const accessToken = await getAccessToken(
      env.msTenantId,
      env.msClientId,
      env.msClientSecret
    );
    console.log("Access token obtained");

    // Step 3: Find target Exchange calendar
    console.log(
      `Finding target calendar: ${env.msTargetCalendarName}`
    );
    const targetCalendar = await findTargetCalendar(
      accessToken,
      env.msUserId,
      env.msTargetCalendarName
    );
    console.log(`Found target calendar: ${targetCalendar.id}`);

    // Step 4: Sync events using update-or-create pattern
    // This preserves manual showAs="free" status changes
    console.log(`Syncing ${iCloudEvents.length} events from iCloud...`);
    const syncResult = await syncEvents(
      accessToken,
      env.msUserId,
      targetCalendar.id,
      iCloudEvents,
      env.timezone,
      window
    );
    console.log(`Sync complete: ${syncResult.created} created, ${syncResult.updated} updated, ${syncResult.skipped} skipped`);

    // Step 5: Delete orphaned events (exist in Exchange but not in iCloud)
    // Only delete orphaned events within the sync window (not old events outside window)
    // IMPORTANT: Only delete events that have iCalUId set (synced events), not manually created ones
    console.log("Checking for orphaned events to delete (within sync window only)...");
    const existingEventsInWindow = await listEventsInWindow(
      accessToken,
      env.msUserId,
      targetCalendar.id,
      window.start,
      window.end
    );
    
    const iCloudUids = new Set(iCloudEvents.map(e => e.uid));
    
    // Only consider events with iCalUId as potential orphans (these are synced events)
    // Events without iCalUId might be manually created and should be preserved
    // Also exclude events we just created/updated in this sync
    const orphanedEvents = existingEventsInWindow.filter(e => {
      // Don't delete events we just created/updated
      if (syncResult.createdEventIds.has(e.id)) {
        return false;
      }
      if (!e.iCalUId) {
        // Event has no iCalUId - might be manually created, don't delete
        return false;
      }
      // Event has iCalUId but doesn't match any iCloud event - it's an orphan
      return !iCloudUids.has(e.iCalUId);
    });
    
    let deletedCount = 0;
    if (orphanedEvents.length > 0) {
      console.log(`Found ${orphanedEvents.length} orphaned event(s) to delete (events with iCalUId that don't match iCloud)`);
      console.log(`Orphaned events: ${orphanedEvents.map(e => `"${e.subject}" (${e.iCalUId})`).join(", ")}`);
      
      for (const event of orphanedEvents) {
        try {
          const deleteUrl = `${GRAPH_BASE_URL}/users/${env.msUserId}/calendars/${targetCalendar.id}/events/${event.id}`;
          const deleteResponse = await fetch(deleteUrl, {
            method: "DELETE",
            headers: {
              Authorization: `Bearer ${accessToken}`,
            },
          });
          if (deleteResponse.ok) {
            deletedCount++;
            console.log(`Deleted orphaned event: "${event.subject}" (${event.iCalUId})`);
          } else {
            console.warn(`Failed to delete orphaned event "${event.subject}" (${event.id}): ${deleteResponse.status}`);
          }
        } catch (error) {
          console.error(`Error deleting orphaned event "${event.subject}":`, error);
        }
      }
      console.log(`Deleted ${deletedCount} orphaned event(s)`);
    } else {
      console.log("No orphaned events found");
    }
    
    // Log events without iCalUId that were preserved
    const eventsWithoutUid = existingEventsInWindow.filter(e => !e.iCalUId);
    if (eventsWithoutUid.length > 0) {
      console.log(`Preserved ${eventsWithoutUid.length} event(s) without iCalUId (likely manually created): ${eventsWithoutUid.map(e => `"${e.subject}"`).join(", ")}`);
    }

    const duration = Date.now() - startTime;

    const result = {
      success: true,
      stats: {
        fetched: iCloudEvents.length,
        created: syncResult.created,
        updated: syncResult.updated,
        skipped: syncResult.skipped,
        deleted: deletedCount,
        durationMs: duration,
      },
      window: {
        start: window.start.toISOString(),
        end: window.end.toISOString(),
      },
    };

    console.log("Sync completed successfully:", result);

    res.status(200).json(result);
  } catch (error) {
    const duration = Date.now() - startTime;
    const errorMessage =
      error instanceof Error ? error.message : "Unknown error";
    console.error("Sync failed:", errorMessage, error);

    res.status(500).json({
      success: false,
      error: errorMessage,
      durationMs: duration,
    });
  }
}


