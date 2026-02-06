/**
 * Vercel serverless function to sync iCloud calendars to Exchange calendar
 * Runs every 15 minutes via Vercel Cron Job
 */

import type { VercelRequest, VercelResponse } from "@vercel/node";
import {
  fetchAllEvents,
  discoverCalendars,
  discoverAllCalendars,
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
const ALTVINA_BLOCK_UID_PREFIX = "altvina-";

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
  msSourceCalendarName?: string; // Optional: Exchange calendar to read work appointments → write busy blocks to iCloud
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
    MS_SOURCE_CALENDAR_NAME: process.env.MS_SOURCE_CALENDAR_NAME, // Optional
    SYNC_LOOKBACK_DAYS: process.env.SYNC_LOOKBACK_DAYS,
    SYNC_LOOKAHEAD_DAYS: process.env.SYNC_LOOKAHEAD_DAYS,
    TIMEZONE: process.env.TIMEZONE,
  };

  const trimEnvVar = (value: string): string => value.trim().replace(/^["']|["']$/g, "");
  const optional = (v: string | undefined) => (v ? trimEnvVar(v) : undefined);

  // ICLOUD_CAL3 and MS_SOURCE_CALENDAR_NAME are optional
  const missing = Object.entries(requiredVars)
    .filter(([key, value]) => !value && key !== "ICLOUD_CAL3" && key !== "MS_SOURCE_CALENDAR_NAME")
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

  return {
    icloudUsername: trimEnvVar(requiredVars.ICLOUD_USERNAME!),
    icloudPassword: trimEnvVar(requiredVars.ICLOUD_APP_PASSWORD!),
    icloudCal1: trimEnvVar(requiredVars.ICLOUD_CAL1!),
    icloudCal2: trimEnvVar(requiredVars.ICLOUD_CAL2!),
    icloudCal3: optional(requiredVars.ICLOUD_CAL3),
    msTenantId: trimEnvVar(requiredVars.MS_TENANT_ID!),
    msClientId: trimEnvVar(requiredVars.MS_CLIENT_ID!),
    msClientSecret: trimEnvVar(requiredVars.MS_CLIENT_SECRET!),
    msUserId: trimEnvVar(requiredVars.MS_USER_ID!),
    msTargetCalendarName: trimEnvVar(requiredVars.MS_TARGET_CALENDAR_NAME!),
    msSourceCalendarName: optional(requiredVars.MS_SOURCE_CALENDAR_NAME),
    syncLookbackDays: lookbackDays,
    syncLookaheadDays: lookaheadDays,
    timezone: trimEnvVar(requiredVars.TIMEZONE!),
  };
}

/**
 * Kill switch: when set (e.g. SYNC_PAUSED=true), the sync does nothing and returns immediately.
 * No reads or writes to any calendar. Set in Vercel env vars to stop all changes until you're ready.
 */
function isSyncPaused(): boolean {
  const v = process.env.SYNC_PAUSED;
  if (!v) return false;
  const s = v.trim().toLowerCase();
  return s === "true" || s === "1" || s === "yes" || s === "on";
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
    if (isSyncPaused()) {
      const duration = Date.now() - startTime;
      console.warn("SYNC_PAUSED is set — skipping all calendar operations. No changes made.");
      res.status(200).json({
        success: true,
        paused: true,
        message: "Calendar sync is paused. No reads or writes were performed. Set SYNC_PAUSED to false (or remove it) to resume.",
        durationMs: duration,
      });
      return;
    }

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
      const publicEventsRaw = await fetchPublicCalendarEvents(
        env.icloudCal3,
        "Personal Calendar",
        window.start,
        window.end
      );
      const newFromPublic = publicEventsRaw.filter((e) => !caldavUids.has(e.uid));
      if (newFromPublic.length < publicEventsRaw.length) {
        console.log(
          `Skipping ${publicEventsRaw.length - newFromPublic.length} public event(s) already from CalDAV (prefer CAL1/CAL2 over CAL3)`
        );
      }
      // Treat public calendar as Jay's Calendar (CAL2) so they show SyncSource CAL2, not CAL3
      const publicAsCal2 = newFromPublic.map((e) => ({ ...e, calendarName: env.icloudCal2 }));
      iCloudEvents.push(...publicAsCal2);
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
    let setIcloudCal2ToExactlyOneOf: string[] | undefined;
    if (!cal2Url) {
      console.warn(
        `CAL2 URL not found; ICLOUD_CAL2="${env.icloudCal2}". Fetching all iCloud calendar names.`
      );
      try {
        const allCals = await discoverAllCalendars(env.icloudUsername, env.icloudPassword);
        setIcloudCal2ToExactlyOneOf = allCals.map((c) => c.name);
        console.log(`All iCloud calendar names: ${setIcloudCal2ToExactlyOneOf.map((n) => `"${n}"`).join(", ")}`);
      } catch (err) {
        console.error("Failed to discover all calendars:", err);
      }
    }

    const iCloudByCalAndUid = new Map<string, (typeof iCloudEvents)[0]>();
    for (const ev of iCloudEvents) {
      const n = normalizeCalName(ev.calendarName);
      if (n === normalizeCalName(env.icloudCal1) || n === normalizeCalName(env.icloudCal2)) {
        iCloudByCalAndUid.set(`${n}\0${ev.uid}`, ev);
      }
    }

    // CAL2 UIDs we consider "synced to Outlook" — exclude Altvina blocks (they are iCloud-only for availability)
    const cal2UidsFromIcloud = new Set(
      iCloudEvents
        .filter(
          (e) =>
            normalizeCalName(e.calendarName) === normalizeCalName(env.icloudCal2) &&
            !e.uid.startsWith(ALTVINA_BLOCK_UID_PREFIX)
        )
        .map((e) => e.uid)
    );

    const outlookToIcloud = {
      created: 0,
      updated: 0,
      deleted: 0,
      skippedNoCal2Url: 0,
      candidates: 0, // Outlook events with no SyncSource (would create on CAL2)
      errors: [] as string[],
    };
    const restoredCal1Uids = new Set<string>();

    for (const outlookEvent of existingEventsInWindow) {
      const meta = parseSyncMeta(outlookEvent.body?.content);
      const { uid: metaUid, syncSource } = meta;

      // No SyncSource = event was created manually in Outlook on Personal Busy; default to adding block on Jay's Calendar (CAL2)
      if (!syncSource) {
        outlookToIcloud.candidates++;
        if (!cal2Url) {
          outlookToIcloud.skippedNoCal2Url++;
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
          outlookToIcloud.created++;
          console.log(`Outlook→iCloud: created "${outlookEvent.subject}" on CAL2 and tagged Outlook (${newUid})`);
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          outlookToIcloud.errors.push(`${outlookEvent.subject}: ${msg}`);
          console.error(`Outlook→iCloud: failed to create "${outlookEvent.subject}" on CAL2:`, err);
        }
        continue;
      }

      if (!metaUid) continue; // need UID for update/delete from iCloud
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
            outlookToIcloud.updated++;
            console.log(`Outlook→iCloud: updated "${outlookEvent.subject}" (${metaUid}) on ${syncSource}`);
          } catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            outlookToIcloud.errors.push(`update ${outlookEvent.subject}: ${msg}`);
            console.error(`Outlook→iCloud: failed to update "${outlookEvent.subject}":`, err);
          }
        }
      } else if (calUrl && !metaUid.startsWith(ALTVINA_BLOCK_UID_PREFIX)) {
        // Event exists on Outlook (synced before) but missing on iCloud — restore to iCloud (e.g. after wrongful delete)
        try {
          const outlookNorm = graphEventToNormalizedEvent(outlookEvent, {
            uid: metaUid,
            calendarName: calName,
          });
          await createIcloudEvent(
            env.icloudUsername,
            env.icloudPassword,
            calUrl,
            { ...outlookNorm, uid: metaUid }
          );
          if (syncSource === "CAL2") cal2UidsFromIcloud.add(metaUid);
          if (syncSource === "CAL1") restoredCal1Uids.add(metaUid);
          outlookToIcloud.created++;
          console.log(`Outlook→iCloud: restored "${outlookEvent.subject}" (${metaUid}) to ${syncSource}`);
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          outlookToIcloud.errors.push(`restore ${outlookEvent.subject}: ${msg}`);
          console.error(`Outlook→iCloud: failed to restore "${outlookEvent.subject}" to ${syncSource}:`, err);
        }
      }
    }

    // Delete from iCloud only when we're sure the user removed the event from Outlook.
    // Use both: (1) body SyncSource+UID, and (2) iCalUId match (in case body was missing/lost).
    const outlookKeys = new Set(
      existingEventsInWindow
        .map((e) => {
          const m = parseSyncMeta(e.body?.content);
          if (m.syncSource && m.uid) return `${m.syncSource}\0${m.uid}`;
          return null;
        })
        .filter((k): k is string => k !== null)
    );
    const outlookUidsFromICal = new Set(
      existingEventsInWindow.map((e) => e.iCalUId).filter((uid): uid is string => Boolean(uid))
    );
    for (const ev of iCloudEvents) {
      if (ev.calendarName !== env.icloudCal1 && ev.calendarName !== env.icloudCal2) continue;
      if (!ev.eventUrl) continue;
      // Altvina busy blocks (uid starts with altvina-) are managed by Exchange→iCloud only; never delete based on Outlook
      if (ev.uid.startsWith(ALTVINA_BLOCK_UID_PREFIX)) continue;
      const tag = ev.calendarName === env.icloudCal1 ? "CAL1" : "CAL2";
      const key = `${tag}\0${ev.uid}`;
      // Do not delete if Outlook has this event (by body tag or by iCalUId)
      if (outlookKeys.has(key) || outlookUidsFromICal.has(ev.uid)) continue;
      try {
        await deleteIcloudEventByUrl(
          env.icloudUsername,
          env.icloudPassword,
          ev.eventUrl
        );
        outlookToIcloud.deleted++;
        console.log(`Outlook→iCloud: deleted "${ev.title}" (${ev.uid}) from ${tag}`);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        outlookToIcloud.errors.push(`delete ${ev.title}: ${msg}`);
        console.error(`Outlook→iCloud: failed to delete "${ev.title}":`, err);
      }
    }

    // Step 5b: Exchange (Altvina) → iCloud: write busy blocks to Jay's Calendar so personal calendar shows work appointments
    let exchangeToIcloudCreated = 0;
    let exchangeToIcloudUpdated = 0;
    let exchangeToIcloudDeleted = 0;
    if (env.msSourceCalendarName && cal2Url) {
      try {
        const sourceCalendar = await findTargetCalendar(
          accessToken,
          env.msUserId,
          env.msSourceCalendarName
        );
        const sourceEvents = await listEventsInWindow(
          accessToken,
          env.msUserId,
          sourceCalendar.id,
          window.start,
          window.end
        );
        const busyEvents = sourceEvents.filter((e) => e.showAs && e.showAs !== "free");
        const busyIds = new Set(busyEvents.map((e) => e.id));
        const altvinaUidToEvent = new Map<string, (typeof iCloudEvents)[0]>();
        for (const ev of iCloudEvents) {
          if (ev.uid.startsWith(ALTVINA_BLOCK_UID_PREFIX) && normalizeCalName(ev.calendarName) === normalizeCalName(env.icloudCal2) && ev.eventUrl) {
            altvinaUidToEvent.set(ev.uid, ev);
          }
        }
        for (const e of busyEvents) {
          const uid = ALTVINA_BLOCK_UID_PREFIX + e.id;
          const start = new Date(e.start.dateTime);
          const end = new Date(e.end.dateTime);
          const normalized = {
            uid,
            title: "Altvina Engagement",
            description: "",
            start,
            end,
            calendarName: env.icloudCal2,
            isAllDay: e.isAllDay ?? false,
          };
          const existing = altvinaUidToEvent.get(uid);
          if (existing?.eventUrl) {
            const same = existing.start.getTime() === start.getTime() && existing.end.getTime() === end.getTime();
            if (!same) {
              await updateIcloudEvent(env.icloudUsername, env.icloudPassword, existing.eventUrl, normalized);
              exchangeToIcloudUpdated++;
            }
          } else {
            await createIcloudEvent(env.icloudUsername, env.icloudPassword, cal2Url, normalized);
            exchangeToIcloudCreated++;
          }
        }
        for (const [uid, ev] of altvinaUidToEvent) {
          const exchangeId = uid.slice(ALTVINA_BLOCK_UID_PREFIX.length);
          if (!busyIds.has(exchangeId)) {
            await deleteIcloudEventByUrl(env.icloudUsername, env.icloudPassword, ev.eventUrl!);
            exchangeToIcloudDeleted++;
          }
        }
        if (exchangeToIcloudCreated || exchangeToIcloudUpdated || exchangeToIcloudDeleted) {
          console.log(`Exchange→iCloud: ${exchangeToIcloudCreated} created, ${exchangeToIcloudUpdated} updated, ${exchangeToIcloudDeleted} deleted (busy blocks on ${env.icloudCal2})`);
        }
      } catch (err) {
        console.error("Exchange→iCloud (Altvina busy blocks):", err);
      }
    }

    // Step 6: iCloud → Outlook sync (with SyncSource in body). Exclude Altvina blocks so they stay iCloud-only.
    const iCloudEventsForOutlook = iCloudEvents.filter((e) => !e.uid.startsWith(ALTVINA_BLOCK_UID_PREFIX));
    const syncOptions = {
      cal1Name: env.icloudCal1,
      cal2Name: env.icloudCal2,
    };
    console.log(`Syncing ${iCloudEventsForOutlook.length} events from iCloud (excluding ${iCloudEvents.length - iCloudEventsForOutlook.length} Altvina blocks)...`);
    const syncResult = await syncEvents(
      accessToken,
      env.msUserId,
      targetCalendar.id,
      iCloudEventsForOutlook,
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
        if (cal1Uids.size === 0 && restoredCal1Uids.size === 0) return false; // No CAL1 data this run; don't delete
        return !cal1Uids.has(m.uid) && !restoredCal1Uids.has(m.uid);
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
      outlookToIcloud: {
        created: outlookToIcloud.created,
        updated: outlookToIcloud.updated,
        deleted: outlookToIcloud.deleted,
        skippedNoCal2Url: outlookToIcloud.skippedNoCal2Url,
        candidates: outlookToIcloud.candidates,
        errors: outlookToIcloud.errors.length > 0 ? outlookToIcloud.errors : undefined,
        setIcloudCal2ToExactlyOneOf,
      },
      /** When MS_SOURCE_CALENDAR_NAME is set: busy blocks on Altvina calendar written to iCloud (Jay's Calendar) */
      exchangeToIcloud: {
        created: exchangeToIcloudCreated,
        updated: exchangeToIcloudUpdated,
        deleted: exchangeToIcloudDeleted,
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


