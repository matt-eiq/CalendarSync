# Google Calendar to Outlook Sync

A Google Apps Script that mirrors events from your primary Google Calendar to Outlook by creating invite copies on a dedicated sync calendar. Events sync automatically every 10 minutes.

## How it works

The script creates a hidden "[Sync] Google -> Outlook" calendar in your Google account. Every 10 minutes, it reads your primary Google Calendar, creates matching events on the sync calendar, and invites your Outlook email as an attendee. Outlook receives the invite and adds it to your calendar.

Events that change or get deleted in Google are automatically updated or cancelled in Outlook. Events where your Outlook email is already an attendee (e.g., cross-org meetings) are skipped to avoid duplicates.

All synced events are prefixed with `[GCal Sync]` in the subject line so you can easily filter them in both Gmail and Outlook.

## Setup

### 1. Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Delete any boilerplate code and paste the contents of `Code.gs`
3. Name the project (e.g., "Calendar Sync")

### 2. Enable the Calendar API

1. In the Apps Script editor, click the **+** next to **Services** in the left sidebar
2. Select **Google Calendar API** and click **Add**

### 3. Configure your Outlook email

In `Code.gs`, find the `CONFIG` object at the top and set your Outlook email:

```js
OUTLOOK_EMAIL: 'you@yourcompany.com',
```

### 4. Run setup

1. In the editor, select the `setup` function from the dropdown at the top
2. Click **Run**
3. Grant the permissions when prompted (calendar access and email sending for invites)

The script will create the sync calendar, install the 10-minute trigger, and run the first sync. You're done with the script side.

### 5. Set up the Gmail filter (required)

When Outlook receives the calendar invite, it may send an acceptance reply back to Gmail. Without a filter, this creates a loop. You need a Gmail filter to catch and discard these replies.

1. In Gmail, click the **search options** icon (down arrow) in the search bar
2. Set the filter criteria:
   - **Subject:** `[GCal Sync]`
   - **From:** your Outlook email address (the one in `CONFIG.OUTLOOK_EMAIL`)
3. Click **Create filter**
4. Check:
   - **Skip the Inbox (Archive it)**
   - **Delete it**
5. Click **Create filter**

This catches the automated reply-backs from Outlook without affecting any other mail.

### 6. Set up the Outlook filter (recommended)

The synced invites arrive in Outlook as normal meeting invitations. To keep them from cluttering your inbox:

1. In Outlook, go to **Settings** (gear icon) > **Mail** > **Rules**
2. Click **Add new rule**
3. Name it something like "GCal Sync Auto-Accept"
4. Set the condition:
   - **Subject contains:** `[GCal Sync]`
5. Set the action(s) you want, for example:
   - **Move to:** a specific folder, or
   - **Delete** the notification email (the calendar event will still appear on your calendar), or
   - **Mark as read**
6. Click **Save**

The `[GCal Sync]` prefix is unique to this script, so the rule won't accidentally catch any other emails.

## Configuration options

All options are in the `CONFIG` object at the top of `Code.gs`:

| Option | Default | Description |
|---|---|---|
| `OUTLOOK_EMAIL` | (must set) | Your Outlook email address |
| `SYNC_WINDOW_DAYS` | `28` | How far ahead to sync events |
| `TRIGGER_INTERVAL_MINUTES` | `10` | How often the sync runs |
| `SUBJECT_TAG` | `[GCal Sync] ` | Prefix on all mirrored event subjects |
| `SYNC_CALENDAR_NAME` | `[Sync] Google -> Outlook` | Name of the sync calendar |

## Uninstalling

Run the `teardown` function from the Apps Script editor. This will:

- Remove the time-driven trigger
- Delete all mirror events (sending cancellations to Outlook)
- Delete the sync calendar
- Clean up all stored script properties

## Troubleshooting

**Duplicate events in Outlook:** Make sure the Gmail filter is set up correctly (step 5). Also check that `CONFIG.OUTLOOK_EMAIL` matches the email that receives the invites.

**Events not appearing in Outlook:** Open the Apps Script editor and check **Executions** in the left sidebar for errors. The most common issue is forgetting to enable the Calendar API service (step 2).

**Sync map corrupted:** Run the `resetSyncMap` function from the editor. The next sync run will rebuild the mapping from the mirror events' metadata.
