/**
 * Vercel serverless function to sync iCloud calendars to Exchange calendar
 * Runs every 15 minutes via Vercel Cron Job
 */

import type { VercelRequest, VercelResponse } from "@vercel/node";
import {
  fetchAllEvents,
  discoverCalendars,
  createIcloudEvent,
  updateIcloudEvent,
  deleteIcloudEventByUrl,
} from "../lib/icloud.js";
import {
  getAccessToken,
  findTargetCalendar,
  syncEvents,
  listEventsInWindow,
  updateEvent,
  parseSyncMeta,
  graphEventToNormalizedEvent,
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
    // If the public URL is a published view of Jay's Calendar (CAL2), the same UIDs appear from both
    // CalDAV and public. We only add public events whose UID is not already in iCloudEvents so we
    // keep the CalDAV source (CAL2 with eventUrl for write-back) and avoid tagging as CAL3.
    const caldavUids = new Set(iCloudEvents.map((e) => e.uid));
    if (env.icloudCal3) {
      console.log(`Fetching events from public calendar URL: ${env.icloudCal3}`);
      const { fetchPublicCalendarEvents } = await import("../lib/icloud.js");
      const publicEvents = await fetchPublicCalendarEvents(
        env.icloudCal3,
        "Personal Calendar",
        window.start,
        window.end
      );
      const newFromPublic = publicEvents.filter((e) => !caldavUids.has(e.uid));
      if (newFromPublic.length < publicEvents.length) {
        console.log(
          `Skipping ${publicEvents.length - newFromPublic.length} public event(s) already from CalDAV (prefer CAL1/CAL2 over CAL3)`
        );
      }
      iCloudEvents.push(...newFromPublic);
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

    // Step 4: Fetch Outlook events with body (for two-way sync: parse SyncSource + UID)
    const existingEventsInWindow = await listEventsInWindow(
      accessToken,
      env.msUserId,
      targetCalendar.id,
      window.start,
      window.end,
      { includeBody: true }
    );

    // Step 5: Outlook → iCloud write-back (create new on CAL2, update/delete by origin)
    const caldavCalendars = await discoverCalendars(
      env.icloudUsername,
      env.icloudPassword,
      [env.icloudCal1, env.icloudCal2]
    );
    // Normalize names for matching (e.g. curly vs straight apostrophe: "Jay's" from iCloud vs env)
    const normalizeCalName = (s: string) =>
      s.trim().replace(/\u2019/g, "'").replace(/\u2018/g, "'");
    const cal1Url = caldavCalendars.find(
      (c) => normalizeCalName(c.name) === normalizeCalName(env.icloudCal1)
    )?.url;
    const cal2Url = caldavCalendars.find(
      (c) => normalizeCalName(c.name) === normalizeCalName(env.icloudCal2)
    )?.url;
    if (!cal2Url) {
      console.warn(
        `CAL2 URL not found; discovered: ${caldavCalendars.map((c) => `"${c.name}"`).join(", ")}. ICLOUD_CAL2="${env.icloudCal2}". Outlook-created events will not write back to iCloud.`
      );
    }

    const iCloudByCalAndUid = new Map<string, (typeof iCloudEvents)[0]>();
    for (const ev of iCloudEvents) {
      const n = normalizeCalName(ev.calendarName);
      if (n === normalizeCalName(env.icloudCal1) || n === normalizeCalName(env.icloudCal2)) {
        iCloudByCalAndUid.set(`${n}\0${ev.uid}`, ev);
      }
    }

    const cal2UidsFromIcloud = new Set(
      iCloudEvents.filter((e) => normalizeCalName(e.calendarName) === normalizeCalName(env.icloudCal2)).map((e) => e.uid)
    );

    for (const outlookEvent of existingEventsInWindow) {
      const meta = parseSyncMeta(outlookEvent.body?.content);
      const { uid: metaUid, syncSource } = meta;

      if (!syncSource || !metaUid) {
        // Created in Outlook → create on iCloud CAL2 (Jay's Calendar) and tag Outlook with SyncSource + UID
        if (!cal2Url) {
          console.warn(`Skipping Outlook-created event "${outlookEvent.subject}" (no CAL2 URL)`);
          continue;
        }
        console.log(`Outlook-created event "${outlookEvent.subject}"; creating on CAL2...`);
        try {
          const normalized = graphEventToNormalizedEvent(outlookEvent, {
            uid: "",
            calendarName: env.icloudCal2,
          });
          const newUid = await createIcloudEvent(
            env.icloudUsername,
            env.icloudPassword,
            cal2Url,
            { ...normalized, uid: normalized.uid || "" }
          );
          cal2UidsFromIcloud.add(newUid);
          const toUpdate = graphEventToNormalizedEvent(outlookEvent, {
            uid: newUid,
            calendarName: env.icloudCal2,
          });
          await updateEvent(
            accessToken,
            env.msUserId,
            targetCalendar.id,
            outlookEvent.id,
            toUpdate,
            env.timezone,
            outlookEvent.showAs === "free" ? "free" : undefined,
            "CAL2"
          );
          console.log(`Outlook→iCloud: created "${outlookEvent.subject}" on CAL2 and tagged Outlook (${newUid})`);
        } catch (err) {
          console.error(`Outlook→iCloud: failed to create "${outlookEvent.subject}" on CAL2:`, err);
        }
        continue;
      }

      if (syncSource === "CAL3") continue; // Public calendar read-only

      const calName = syncSource === "CAL1" ? env.icloudCal1 : env.icloudCal2;
      const calUrl = syncSource === "CAL1" ? cal1Url : cal2Url;
      const iCloudEv = iCloudByCalAndUid.get(`${normalizeCalName(calName)}\0${metaUid}`);

      if (iCloudEv?.eventUrl) {
        const outlookNorm = graphEventToNormalizedEvent(outlookEvent, { calendarName: calName });
        const same =
          outlookNorm.title === iCloudEv.title &&
          outlookNorm.start.getTime() === iCloudEv.start.getTime() &&
          outlookNorm.end.getTime() === iCloudEv.end.getTime();
        if (!same) {
          try {
            await updateIcloudEvent(
              env.icloudUsername,
              env.icloudPassword,
              iCloudEv.eventUrl,
              { ...outlookNorm, uid: metaUid }
            );
            console.log(`Outlook→iCloud: updated "${outlookEvent.subject}" (${metaUid}) on ${syncSource}`);
          } catch (err) {
            console.error(`Outlook→iCloud: failed to update "${outlookEvent.subject}":`, err);
          }
        }
      }
    }

    // Delete from iCloud when user deleted the event in Outlook (no Outlook event with this SyncSource+UID)
    const outlookKeys = new Set(
      existingEventsInWindow
        .map((e) => {
          const m = parseSyncMeta(e.body?.content);
          if (m.syncSource && m.uid) return `${m.syncSource}\0${m.uid}`;
          return null;
        })
        .filter((k): k is string => k !== null)
    );
    for (const ev of iCloudEvents) {
      if (ev.calendarName !== env.icloudCal1 && ev.calendarName !== env.icloudCal2) continue;
      if (!ev.eventUrl) continue;
      const tag = ev.calendarName === env.icloudCal1 ? "CAL1" : "CAL2";
      const key = `${tag}\0${ev.uid}`;
      if (!outlookKeys.has(key)) {
        try {
          await deleteIcloudEventByUrl(
            env.icloudUsername,
            env.icloudPassword,
            ev.eventUrl
          );
          console.log(`Outlook→iCloud: deleted "${ev.title}" (${ev.uid}) from ${tag}`);
        } catch (err) {
          console.error(`Outlook→iCloud: failed to delete "${ev.title}":`, err);
        }
      }
    }

    // Step 6: iCloud → Outlook sync (with SyncSource in body)
    const syncOptions = {
      cal1Name: env.icloudCal1,
      cal2Name: env.icloudCal2,
    };
    console.log(`Syncing ${iCloudEvents.length} events from iCloud...`);
    const syncResult = await syncEvents(
      accessToken,
      env.msUserId,
      targetCalendar.id,
      iCloudEvents,
      env.timezone,
      window,
      syncOptions
    );
    console.log(`Sync complete: ${syncResult.created} created, ${syncResult.updated} updated, ${syncResult.skipped} skipped`);

    // Step 7: Orphan delete (Outlook events whose SyncSource+UID no longer exists on iCloud)
    // Only delete when we have a non-empty UID set for that source; otherwise we may have failed
    // to fetch that calendar (e.g. name mismatch) and would wrongly delete valid events.
    const cal1Uids = new Set(
      iCloudEvents.filter((e) => normalizeCalName(e.calendarName) === normalizeCalName(env.icloudCal1)).map((e) => e.uid)
    );
    const cal3Uids = new Set(
      iCloudEvents.filter(
        (e) =>
          normalizeCalName(e.calendarName) !== normalizeCalName(env.icloudCal1) &&
          normalizeCalName(e.calendarName) !== normalizeCalName(env.icloudCal2)
      ).map((e) => e.uid)
    );

    const orphanedEvents = existingEventsInWindow.filter((e) => {
      if (syncResult.createdEventIds.has(e.id)) return false;
      const m = parseSyncMeta(e.body?.content);
      if (!m.syncSource || !m.uid) return false;
      if (m.syncSource === "CAL1") {
        if (cal1Uids.size === 0) return false; // No CAL1 data this run; don't delete
        return !cal1Uids.has(m.uid);
      }
      if (m.syncSource === "CAL2") {
        if (cal2UidsFromIcloud.size === 0) return false;
        return !cal2UidsFromIcloud.has(m.uid);
      }
      if (m.syncSource === "CAL3") {
        if (cal3Uids.size === 0) return false;
        return !cal3Uids.has(m.uid);
      }
      return false;
    });

    let deletedCount = 0;
    if (orphanedEvents.length > 0) {
      console.log(`Found ${orphanedEvents.length} orphaned event(s) to delete`);
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


