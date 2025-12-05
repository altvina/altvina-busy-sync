/**
 * Vercel serverless function to sync iCloud calendars to Exchange calendar
 * Runs every 15 minutes via Vercel Cron Job
 */

import type { VercelRequest, VercelResponse } from "@vercel/node";
import { fetchAllEvents } from "../lib/icloud.js";
import {
  getAccessToken,
  findTargetCalendar,
  deleteEventsInWindow,
  createEvents,
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
 * Calculate extended delete window to catch old test events
 * Deletes events from further back to ensure clean slate
 */
function calculateDeleteWindow(
  lookbackDays: number,
  lookaheadDays: number
): SyncWindow {
  const now = new Date();
  // Delete from 30 days ago to catch old test events
  const start = new Date(now);
  start.setDate(start.getDate() - 30);
  start.setHours(0, 0, 0, 0);

  // Delete up to the same lookahead
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
    MS_TENANT_ID: process.env.MS_TENANT_ID,
    MS_CLIENT_ID: process.env.MS_CLIENT_ID,
    MS_CLIENT_SECRET: process.env.MS_CLIENT_SECRET,
    MS_USER_ID: process.env.MS_USER_ID,
    MS_TARGET_CALENDAR_NAME: process.env.MS_TARGET_CALENDAR_NAME,
    SYNC_LOOKBACK_DAYS: process.env.SYNC_LOOKBACK_DAYS,
    SYNC_LOOKAHEAD_DAYS: process.env.SYNC_LOOKAHEAD_DAYS,
    TIMEZONE: process.env.TIMEZONE,
  };

  const missing = Object.entries(requiredVars)
    .filter(([_, value]) => !value)
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
    console.log(
      `Fetching events from iCloud calendars: ${env.icloudCal1}, ${env.icloudCal2}`
    );
    const iCloudEvents = await fetchAllEvents(
      env.icloudUsername,
      env.icloudPassword,
      [env.icloudCal1, env.icloudCal2],
      window.start,
      window.end
    );
    console.log(`Fetched ${iCloudEvents.length} events from iCloud`);

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

    // Step 4: Delete all existing events in an extended window
    // Use extended window to catch old test events outside sync range
    const deleteWindow = calculateDeleteWindow(
      env.syncLookbackDays,
      env.syncLookaheadDays
    );
    console.log(`Deleting existing events in extended window: ${deleteWindow.start.toISOString()} to ${deleteWindow.end.toISOString()}`);
    const deletedCount = await deleteEventsInWindow(
      accessToken,
      env.msUserId,
      targetCalendar.id,
      deleteWindow.start,
      deleteWindow.end,
      env.timezone
    );
    console.log(`Deleted ${deletedCount} existing events`);

    // Step 5: Create new events from iCloud
    console.log(`Creating ${iCloudEvents.length} new events...`);
    const createdCount = await createEvents(
      accessToken,
      env.msUserId,
      targetCalendar.id,
      iCloudEvents,
      env.timezone
    );
    console.log(`Created ${createdCount} new events`);

    const duration = Date.now() - startTime;

    const result = {
      success: true,
      stats: {
        fetched: iCloudEvents.length,
        deleted: deletedCount,
        created: createdCount,
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

