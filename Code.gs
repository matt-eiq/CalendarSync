// =============================================================================
// Google Calendar → Outlook Sync
//
// Mirrors events from your primary Google Calendar to a dedicated sync calendar,
// inviting your Outlook email address so events appear on your Outlook calendar.
//
// Setup:
//   1. Enable the Advanced Calendar Service (Calendar API v3) in the Apps Script
//      editor: click "+" next to Services → select "Google Calendar API" → Add.
//   2. Set OUTLOOK_EMAIL below to your Outlook email address.
//   3. Run setup() once from the editor to create the sync calendar and trigger.
//   4. Set up a Gmail filter to prevent forwarded-back invites from creating
//      duplicates (see README / plan for details).
// =============================================================================

// ── Configuration ────────────────────────────────────────────────────────────

const CONFIG = {
  // ** SET THIS ** to your Outlook email address
  OUTLOOK_EMAIL: 'YOUR_OUTLOOK_EMAIL@example.com',

  // Calendar ID of the source calendar (usually 'primary')
  SOURCE_CALENDAR_ID: 'primary',

  // Name of the dedicated sync calendar (created automatically by setup())
  SYNC_CALENDAR_NAME: '[Sync] Google → Outlook',

  // How far ahead to sync, in days
  SYNC_WINDOW_DAYS: 28,

  // Trigger interval in minutes
  TRIGGER_INTERVAL_MINUTES: 120,

  // PropertiesService key names
  PROP_SYNC_CALENDAR_ID: 'SYNC_CALENDAR_ID',
  PROP_SYNC_MAP_PREFIX: 'SYNC_MAP_',
  PROP_SYNC_MAP_COUNT: 'SYNC_MAP_COUNT',

  // Prefix added to every mirror event subject — use this in your Outlook
  // rule to filter all synced invites (e.g. "Subject contains [GCal Sync]")
  SUBJECT_TAG: '[GCal Sync] ',

  // Extended property key stored on mirror events (for orphan recovery)
  EXTENDED_PROP_KEY: 'syncSourceId',

  // Max bytes per PropertiesService value (leaving headroom under the 9KB limit)
  CHUNK_SIZE: 8000,

  // Max script execution time before saving partial progress (5 min in ms)
  MAX_RUNTIME_MS: 300000,

  // Max results per Calendar API page
  PAGE_SIZE: 250,
};

// ── Utility Functions ────────────────────────────────────────────────────────

/**
 * Compute an MD5 hash of the sync-relevant fields of a Calendar event.
 * Used for change detection — we only update mirror events when this changes.
 */
function computeHash(event) {
  const parts = [
    event.summary || '',
    event.start.dateTime || event.start.date || '',
    event.end.dateTime || event.end.date || '',
    event.location || '',
    event.status || '',
    event.transparency || '',
  ];
  const str = parts.join('\x1f');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return digest.map(function(b) {
    return ('0' + ((b + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

/**
 * Determine whether a source event should be skipped (not mirrored).
 *
 * Skip if:
 *   - The user has declined the event
 *   - The Outlook email is already an attendee (event is already visible to Outlook)
 *   - The event is cancelled
 */
function shouldSkipEvent(event, userEmail) {
  // Skip cancelled events
  if (event.status === 'cancelled') {
    return true;
  }

  const attendees = event.attendees || [];
  const outlookLower = CONFIG.OUTLOOK_EMAIL.toLowerCase();

  for (var i = 0; i < attendees.length; i++) {
    const attendee = attendees[i];
    const email = (attendee.email || '').toLowerCase();

    // Skip if the user has declined
    if (email === userEmail.toLowerCase() && attendee.responseStatus === 'declined') {
      return true;
    }

    // Skip if Outlook email is already an attendee (Outlook already knows about this event)
    if (email === outlookLower) {
      return true;
    }
  }

  return false;
}

/**
 * Build the Calendar API v3 Event resource for a mirror event.
 */
function buildMirrorPayload(sourceEvent) {
  const payload = {
    summary: CONFIG.SUBJECT_TAG + (sourceEvent.summary || '(No title)'),
    start: sourceEvent.start,
    end: sourceEvent.end,
    attendees: [
      { email: CONFIG.OUTLOOK_EMAIL }
    ],
    reminders: {
      useDefault: false,
      overrides: []
    },
    transparency: sourceEvent.transparency || 'opaque',
    status: 'confirmed',
    extendedProperties: {
      private: {}
    }
  };

  if (sourceEvent.location) {
    payload.location = sourceEvent.location;
  }

  payload.extendedProperties.private[CONFIG.EXTENDED_PROP_KEY] = sourceEvent.id;

  return payload;
}

// ── Sync Map Persistence ─────────────────────────────────────────────────────

/**
 * Load the sync map from PropertiesService.
 * The map is stored as chunked JSON to handle the 9KB per-property limit.
 *
 * Returns: { sourceEventId: { mirrorId: string, hash: string } }
 */
function loadSyncMap() {
  const props = PropertiesService.getScriptProperties();
  const countStr = props.getProperty(CONFIG.PROP_SYNC_MAP_COUNT);

  if (!countStr) {
    return {};
  }

  const count = parseInt(countStr, 10);
  let json = '';

  for (var i = 0; i < count; i++) {
    const chunk = props.getProperty(CONFIG.PROP_SYNC_MAP_PREFIX + i);
    if (chunk) {
      json += chunk;
    }
  }

  try {
    return JSON.parse(json);
  } catch (e) {
    console.error('Failed to parse sync map, resetting: ' + e.message);
    return {};
  }
}

/**
 * Save the sync map to PropertiesService, chunking if necessary.
 */
function saveSyncMap(syncMap) {
  const props = PropertiesService.getScriptProperties();
  const json = JSON.stringify(syncMap);

  // Clear old chunks first
  const oldCountStr = props.getProperty(CONFIG.PROP_SYNC_MAP_COUNT);
  if (oldCountStr) {
    const oldCount = parseInt(oldCountStr, 10);
    for (var i = 0; i < oldCount; i++) {
      props.deleteProperty(CONFIG.PROP_SYNC_MAP_PREFIX + i);
    }
  }

  // Write new chunks
  const chunks = [];
  for (var start = 0; start < json.length; start += CONFIG.CHUNK_SIZE) {
    chunks.push(json.substring(start, start + CONFIG.CHUNK_SIZE));
  }

  // Handle empty map
  if (chunks.length === 0) {
    chunks.push('{}');
  }

  const batch = {};
  batch[CONFIG.PROP_SYNC_MAP_COUNT] = String(chunks.length);
  for (var j = 0; j < chunks.length; j++) {
    batch[CONFIG.PROP_SYNC_MAP_PREFIX + j] = chunks[j];
  }

  props.setProperties(batch);
}

// ── Calendar Management ──────────────────────────────────────────────────────

/**
 * Get the sync calendar ID from properties, or find/create the sync calendar.
 */
function getOrCreateSyncCalendar() {
  const props = PropertiesService.getScriptProperties();
  let calId = props.getProperty(CONFIG.PROP_SYNC_CALENDAR_ID);

  // Verify the stored calendar still exists
  if (calId) {
    try {
      Calendar.Calendars.get(calId);
      return calId;
    } catch (e) {
      // Calendar was deleted externally, recreate it
      console.log('Stored sync calendar not found, will create a new one.');
      calId = null;
    }
  }

  // Search for an existing calendar with the sync name
  const calList = Calendar.CalendarList.list();
  const items = calList.items || [];
  for (var i = 0; i < items.length; i++) {
    if (items[i].summary === CONFIG.SYNC_CALENDAR_NAME) {
      calId = items[i].id;
      props.setProperty(CONFIG.PROP_SYNC_CALENDAR_ID, calId);
      console.log('Found existing sync calendar: ' + calId);
      return calId;
    }
  }

  // Create a new sync calendar
  const newCal = Calendar.Calendars.insert({
    summary: CONFIG.SYNC_CALENDAR_NAME,
    description: 'Auto-generated by Google→Outlook Calendar Sync. Do not manually edit events here.',
    timeZone: Calendar.Settings.get('timezone').value
  });

  calId = newCal.id;
  props.setProperty(CONFIG.PROP_SYNC_CALENDAR_ID, calId);
  console.log('Created sync calendar: ' + calId);
  return calId;
}

/**
 * Fetch all source events from the primary calendar within the sync window.
 * Recurring events are expanded into individual instances via singleEvents=true.
 */
function fetchSourceEvents() {
  const now = new Date();
  const endDate = new Date(now.getTime() + CONFIG.SYNC_WINDOW_DAYS * 24 * 60 * 60 * 1000);

  let allEvents = [];
  let pageToken = null;

  do {
    const params = {
      timeMin: now.toISOString(),
      timeMax: endDate.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
      maxResults: CONFIG.PAGE_SIZE,
      showDeleted: false,
    };

    if (pageToken) {
      params.pageToken = pageToken;
    }

    const response = Calendar.Events.list(CONFIG.SOURCE_CALENDAR_ID, params);
    const items = response.items || [];
    allEvents = allEvents.concat(items);
    pageToken = response.nextPageToken;
  } while (pageToken);

  return allEvents;
}

// ── Mirror Event CRUD ────────────────────────────────────────────────────────

/**
 * Create a new mirror event on the sync calendar.
 */
function createMirrorEvent(sourceEvent, syncCalendarId, syncMap) {
  const payload = buildMirrorPayload(sourceEvent);

  try {
    const mirrorEvent = Calendar.Events.insert(payload, syncCalendarId, {
      sendUpdates: 'all'
    });

    syncMap[sourceEvent.id] = {
      mirrorId: mirrorEvent.id,
      hash: computeHash(sourceEvent),
      endTime: sourceEvent.end.dateTime || sourceEvent.end.date || ''
    };

    console.log('Created mirror: "' + (sourceEvent.summary || '(No title)') + '" → ' + mirrorEvent.id);
    return true;
  } catch (e) {
    console.error('Failed to create mirror for "' + (sourceEvent.summary || '') + '": ' + e.message);
    return false;
  }
}

/**
 * Update an existing mirror event to reflect changes in the source event.
 */
function updateMirrorEvent(sourceEvent, mirrorId, syncCalendarId, syncMap) {
  const payload = buildMirrorPayload(sourceEvent);

  try {
    Calendar.Events.patch(payload, syncCalendarId, mirrorId, {
      sendUpdates: 'all'
    });

    syncMap[sourceEvent.id] = {
      mirrorId: mirrorId,
      hash: computeHash(sourceEvent),
      endTime: sourceEvent.end.dateTime || sourceEvent.end.date || ''
    };

    console.log('Updated mirror: "' + (sourceEvent.summary || '(No title)') + '" (' + mirrorId + ')');
    return true;
  } catch (e) {
    // If the mirror event was deleted externally, remove from map and recreate
    if (e.message && (e.message.indexOf('Not Found') !== -1 || e.message.indexOf('deleted') !== -1)) {
      console.log('Mirror event gone, will recreate: ' + mirrorId);
      delete syncMap[sourceEvent.id];
      return createMirrorEvent(sourceEvent, syncCalendarId, syncMap);
    }
    console.error('Failed to update mirror ' + mirrorId + ': ' + e.message);
    return false;
  }
}

/**
 * Delete a mirror event and send a cancellation to Outlook.
 */
function deleteMirrorEvent(sourceId, mirrorId, syncCalendarId, syncMap) {
  try {
    Calendar.Events.remove(syncCalendarId, mirrorId, {
      sendUpdates: 'all'
    });
    console.log('Deleted mirror: ' + mirrorId + ' (source: ' + sourceId + ')');
  } catch (e) {
    // 404/410 means already deleted — that's fine
    if (e.message && (e.message.indexOf('Not Found') === -1 && e.message.indexOf('Resource has been deleted') === -1)) {
      console.error('Failed to delete mirror ' + mirrorId + ': ' + e.message);
    }
  }

  delete syncMap[sourceId];
}

// ── Orphan Recovery ──────────────────────────────────────────────────────────

/**
 * If the sync map is empty but the sync calendar has events, rebuild the map
 * from the mirror events' extended properties. This prevents duplicate creation
 * if PropertiesService was cleared (e.g., by a code redeploy).
 */
function recoverSyncMap(syncCalendarId) {
  console.log('Sync map is empty but sync calendar may have events. Attempting recovery...');

  const now = new Date();
  const endDate = new Date(now.getTime() + CONFIG.SYNC_WINDOW_DAYS * 24 * 60 * 60 * 1000);

  let recoveredMap = {};
  let pageToken = null;

  do {
    const params = {
      timeMin: now.toISOString(),
      timeMax: endDate.toISOString(),
      singleEvents: true,
      maxResults: CONFIG.PAGE_SIZE,
      showDeleted: false,
    };

    if (pageToken) {
      params.pageToken = pageToken;
    }

    const response = Calendar.Events.list(syncCalendarId, params);
    const items = response.items || [];

    for (var i = 0; i < items.length; i++) {
      const mirrorEvent = items[i];
      const extProps = mirrorEvent.extendedProperties;

      if (extProps && extProps.private && extProps.private[CONFIG.EXTENDED_PROP_KEY]) {
        const sourceId = extProps.private[CONFIG.EXTENDED_PROP_KEY];
        recoveredMap[sourceId] = {
          mirrorId: mirrorEvent.id,
          hash: '', // Force re-check on next reconcile by using empty hash
          endTime: (mirrorEvent.end && (mirrorEvent.end.dateTime || mirrorEvent.end.date)) || ''
        };
      }
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  const count = Object.keys(recoveredMap).length;
  if (count > 0) {
    console.log('Recovered ' + count + ' mirror event mappings from extended properties.');
  } else {
    console.log('No mirror events found on sync calendar. Starting fresh.');
  }

  return recoveredMap;
}

// ── Core Sync Logic ──────────────────────────────────────────────────────────

/**
 * Reconcile source events against the sync map. Creates, updates, and deletes
 * mirror events as needed.
 */
function reconcile(sourceEvents, syncMap, syncCalendarId, userEmail) {
  const startTime = Date.now();
  const seenSourceIds = {};
  let created = 0, updated = 0, deleted = 0, skipped = 0;

  // Process each source event
  for (var i = 0; i < sourceEvents.length; i++) {
    // Execution time guard
    if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
      console.warn('Approaching execution time limit. Saving progress and exiting. ' +
                    'Remaining events will be processed on the next run.');
      break;
    }

    const event = sourceEvents[i];
    seenSourceIds[event.id] = true;

    if (shouldSkipEvent(event, userEmail)) {
      skipped++;
      continue;
    }

    const hash = computeHash(event);
    const existing = syncMap[event.id];

    if (!existing) {
      // New event — create mirror
      if (createMirrorEvent(event, syncCalendarId, syncMap)) {
        created++;
      }
    } else if (existing.hash !== hash) {
      // Changed event — update mirror
      if (updateMirrorEvent(event, existing.mirrorId, syncCalendarId, syncMap)) {
        updated++;
      }
    }
    // else: unchanged, no action needed
  }

  // Delete orphaned mirrors (source event was deleted or fell out of sync window)
  const orphanIds = Object.keys(syncMap).filter(function(sourceId) {
    return !seenSourceIds[sourceId];
  });

  const now = new Date();
  for (var j = 0; j < orphanIds.length; j++) {
    if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
      console.warn('Approaching time limit during orphan cleanup. Will continue next run.');
      break;
    }

    const sourceId = orphanIds[j];
    const entry = syncMap[sourceId];

    // If the event has already ended, it just aged out of the sync window —
    // silently remove it from the map without deleting/cancelling the mirror.
    if (entry.endTime && new Date(entry.endTime) < now) {
      console.log('Event aged out (past end time), removing from map: ' + sourceId);
      delete syncMap[sourceId];
      continue;
    }

    // Future event missing from source — it was genuinely deleted/cancelled.
    deleteMirrorEvent(sourceId, entry.mirrorId, syncCalendarId, syncMap);
    deleted++;
  }

  console.log('Sync complete: ' + created + ' created, ' + updated + ' updated, ' +
              deleted + ' deleted, ' + skipped + ' skipped. ' +
              sourceEvents.length + ' source events processed.');
}

/**
 * Main entry point — called by the time-driven trigger every 10 minutes.
 */
function main() {
  const runStart = Date.now();

  try {
    // 1. Ensure sync calendar exists
    const syncCalendarId = getOrCreateSyncCalendar();

    // 2. Get the user's email for decline detection
    const userEmail = Session.getActiveUser().getEmail();

    // 3. Fetch source events
    const sourceEvents = fetchSourceEvents();
    console.log('Fetched ' + sourceEvents.length + ' source events from primary calendar.');

    // 4. Load sync map (with orphan recovery if needed)
    let syncMap = loadSyncMap();
    if (Object.keys(syncMap).length === 0) {
      syncMap = recoverSyncMap(syncCalendarId);
    }

    // 5. Reconcile
    reconcile(sourceEvents, syncMap, syncCalendarId, userEmail);

    // 6. Persist updated sync map
    saveSyncMap(syncMap);

    console.log('Total runtime: ' + ((Date.now() - runStart) / 1000).toFixed(1) + 's');

  } catch (e) {
    console.error('Sync failed: ' + e.message + '\n' + e.stack);
  }
}

// ── Setup & Teardown ─────────────────────────────────────────────────────────

/**
 * One-time setup. Run this manually from the Apps Script editor.
 * Creates the sync calendar, installs the time-driven trigger, and runs
 * an initial sync.
 */
function setup() {
  // Validate configuration
  if (CONFIG.OUTLOOK_EMAIL === 'YOUR_OUTLOOK_EMAIL@example.com') {
    throw new Error(
      'Please set CONFIG.OUTLOOK_EMAIL to your actual Outlook email address before running setup.'
    );
  }

  console.log('=== Google → Outlook Calendar Sync: Setup ===');

  // Create or find the sync calendar
  const syncCalendarId = getOrCreateSyncCalendar();
  console.log('Sync calendar ready: ' + syncCalendarId);

  // Install the time-driven trigger (if not already installed)
  installTrigger();

  // Run initial sync
  console.log('Running initial sync...');
  main();

  console.log('=== Setup complete! ===');
  console.log('The sync will run every ' + CONFIG.TRIGGER_INTERVAL_MINUTES + ' minutes.');
  console.log('');
  console.log('IMPORTANT: Set up a Gmail filter to prevent forwarded-back invite loops.');
  console.log('Filter criteria:');
  console.log('  - Subject contains: "[Sync] Google"');
  console.log('  - Action: Skip Inbox, Delete');
}

/**
 * Remove all triggers, delete all mirror events (sending cancellations to
 * Outlook), and clean up properties. Run this to fully uninstall the sync.
 */
function teardown() {
  console.log('=== Google → Outlook Calendar Sync: Teardown ===');

  // Remove triggers
  removeTriggers();

  // Delete mirror events (sends cancellations to Outlook)
  const props = PropertiesService.getScriptProperties();
  const syncCalendarId = props.getProperty(CONFIG.PROP_SYNC_CALENDAR_ID);

  if (syncCalendarId) {
    purgeMirrorEvents(syncCalendarId);

    // Optionally delete the sync calendar itself
    try {
      Calendar.Calendars.remove(syncCalendarId);
      console.log('Deleted sync calendar: ' + syncCalendarId);
    } catch (e) {
      console.log('Could not delete sync calendar (may already be gone): ' + e.message);
    }
  }

  // Clear all script properties
  props.deleteAllProperties();

  console.log('=== Teardown complete. All mirror events cancelled and cleaned up. ===');
}

/**
 * Install a time-driven trigger for main() if one doesn't already exist.
 */
function installTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'main') {
      console.log('Trigger for main() already exists. Skipping installation.');
      return;
    }
  }

  ScriptApp.newTrigger('main')
    .timeBased()
    .everyMinutes(CONFIG.TRIGGER_INTERVAL_MINUTES)
    .create();

  console.log('Installed time-driven trigger: main() every ' +
              CONFIG.TRIGGER_INTERVAL_MINUTES + ' minutes.');
}

/**
 * Remove all triggers associated with this script project.
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  console.log('Removed ' + triggers.length + ' trigger(s).');
}

/**
 * Delete all events on the sync calendar, sending cancellation emails.
 * Used during teardown to clean up Outlook.
 */
function purgeMirrorEvents(syncCalendarId) {
  console.log('Purging all mirror events from sync calendar...');
  let pageToken = null;
  let count = 0;

  do {
    const params = {
      maxResults: CONFIG.PAGE_SIZE,
      showDeleted: false,
    };

    if (pageToken) {
      params.pageToken = pageToken;
    }

    let response;
    try {
      response = Calendar.Events.list(syncCalendarId, params);
    } catch (e) {
      console.error('Could not list sync calendar events: ' + e.message);
      return;
    }

    const items = response.items || [];
    for (var i = 0; i < items.length; i++) {
      try {
        Calendar.Events.remove(syncCalendarId, items[i].id, {
          sendUpdates: 'all'
        });
        count++;
      } catch (e) {
        console.error('Failed to delete mirror event ' + items[i].id + ': ' + e.message);
      }
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  console.log('Purged ' + count + ' mirror event(s).');
}

/**
 * Manual reset: clears the sync map without deleting mirror events.
 * Useful if the map gets corrupted. The next run of main() will use
 * orphan recovery to rebuild the map from extended properties.
 */
function resetSyncMap() {
  const props = PropertiesService.getScriptProperties();
  const oldCountStr = props.getProperty(CONFIG.PROP_SYNC_MAP_COUNT);
  if (oldCountStr) {
    const oldCount = parseInt(oldCountStr, 10);
    for (var i = 0; i < oldCount; i++) {
      props.deleteProperty(CONFIG.PROP_SYNC_MAP_PREFIX + i);
    }
    props.deleteProperty(CONFIG.PROP_SYNC_MAP_COUNT);
  }
  console.log('Sync map cleared. Next run will rebuild from extended properties.');
}
