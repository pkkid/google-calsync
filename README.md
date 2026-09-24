# Microsoft calendar → Google Calendar “Nasuni”

A standalone Google Apps Script that mirrors your work calendar into a dedicated Google calendar every five minutes. Microsoft is read-only. The rolling window is **72 hours ago through 336 hours from now**, including events that overlap the window.

Canceled events and events you declined are excluded. Recurring meetings are copied as individual occurrences. Changes, deletions, and events aging out of the window are reconciled on the next successful run.

**Default `mirror` mode removes every event on Nasuni that is absent from the source window, including old and manually added events.** Set `SYNC_MODE=managed` to preserve events this script did not create. The script starts in dry-run mode. Nothing is deployed or scheduled by installing this repository.

## 1. Connect this repo to Apps Script

Use Node.js 22 or newer:

```bash
npm ci
npm run check
npm test
npm run login
```

Enable the Apps Script API in [your Apps Script user settings](https://script.google.com/home/usersettings). Create a new standalone project at [script.google.com](https://script.google.com/), then copy its **Script ID** from Project Settings.

```bash
cp .clasp.example.json .clasp.json
```

Replace the placeholder in `.clasp.json` with that Script ID, leaving `rootDir` as `src`. This local binding is ignored by Git. Then:

```bash
npm run status
npm run push
npm run open
```

The status output should list only the eight JavaScript files and `appsscript.json`. Pushing runs local checks/tests before uploading. The manifest declares the Calendar service and pins the OAuth2 library to version 43. If you use a standard Google Cloud project, also enable the Google Calendar API in that project's Cloud Console. With the default Apps Script Cloud project, adding the Calendar service enables its API automatically. [Google service setup](https://developers.google.com/apps-script/guides/services/advanced)

You do **not** need to deploy a public web app. All functions below are run from the Apps Script editor's function dropdown. Use the same Google account for authorization, initial sync, and trigger installation. Later code changes need only `npm run push`; Script Properties and the installed trigger remain in place. Keep source edits in this repository. [clasp documentation](https://developers.google.com/apps-script/guides/clasp)

## 2. Register the Microsoft app

In the [Microsoft Entra admin center](https://entra.microsoft.com/), open **App registrations → New registration**:

1. Name it something like `Nasuni Calendar Sync` and select accounts in your work directory only.
2. Add a **Web** redirect URI: `https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback`.
3. Under API permissions, add Microsoft Graph **delegated** `Calendars.Read` and `offline_access`. Remove any unneeded default `User.Read` permission. No Microsoft write permissions are used.
4. Under Certificates & secrets, create a client secret. Copy its **value**, not its secret ID, and record its expiration date.
5. Note the **Directory (tenant) ID** and **Application (client) ID**.

If app registration or consent is restricted, IT must register/approve the app. This implementation uses delegated authorization; it does not bypass workplace restrictions. [Microsoft authorization-code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)

## 3. Configure secrets and calendars

In Apps Script **Project Settings → Script Properties**, enter:

| Property | Value |
| --- | --- |
| `MS_TENANT_ID` | Your work tenant ID |
| `MS_CLIENT_ID` | Entra application/client ID |
| `MS_CLIENT_SECRET` | Client secret value |
| `DRY_RUN` | `true` initially; `false` enables writes |
| `SYNC_MODE` | `mirror` (default) or `managed` |
| `GOOGLE_CALENDAR_NAME` | `Nasuni` by default; optional override for a test calendar |
| `GOOGLE_CALENDAR_ID` | Optional initially; setup can resolve and save it by name |
| `MS_CALENDAR_ID` | Optional; omit for the work account's default calendar |

Create a secondary Google calendar named **Nasuni**, if it does not exist. Do not use your primary calendar. For a test run, create a separate secondary calendar and set `GOOGLE_CALENDAR_NAME` to its name. Calendar settings show its ID under **Integrate calendar**.

Run these functions in the Apps Script editor:

1. `configureTargetCalendar` — approve Google's requested access. Resolves one writable secondary calendar by name and saves its ID. If names are ambiguous, enter the intended `GOOGLE_CALENDAR_ID` yourself and rerun it.
2. `authorizeMicrosoft` — open the temporary authorization URL in the execution log and sign in to your **work account**. Complete consent and wait for the callback success page.
3. `previewSync` — reads both calendars and logs create/update/delete/unchanged counts and the exact date window. It never writes calendar events, even when `DRY_RUN=false`.

Names are validated alongside the pinned Google ID on every run. If you rename the calendar, update `GOOGLE_CALENDAR_NAME`. Changing test/production targets requires updating **both** target properties, or deleting the ID and rerunning setup.

Secrets stay in Apps Script. The OAuth2 library stores access/refresh tokens in the authorizing user's **User Properties** and refreshes them automatically. These values are separate from source uploads and are never committed. Keep the Apps Script project private: an editor can access its credentials. Never paste actual credentials into source, `.clasp.json`, a commit, or an issue. Local clasp login credentials are ignored as an additional precaution. [Properties Service](https://developers.google.com/apps-script/guides/properties)

## 4. Enable the sync

After reviewing the preview, preferably testing with a separate calendar:

1. Set `DRY_RUN=false` in Script Properties.
2. Run `syncCalendar` once and inspect the destination calendar.
3. Run `installTrigger`. It verifies read access and replaces this account's existing sync triggers with one five-minute trigger.

Google schedules nominal five-minute intervals; exact timing can vary. Triggers execute under their creator's account. Only one Google account should install the trigger. [Installable triggers](https://developers.google.com/apps-script/guides/triggers/installable)

Run `removeTrigger` to stop scheduled syncs. Alternatively set `DRY_RUN=true` to leave the schedule running as previews. Neither action deletes calendar data.

## Behavior and recovery

- Fetches every Microsoft and Google page before writing. Source read/authentication failures cause no calendar writes; failed creates/updates prevent the deletion phase.
- A successfully fetched **empty** Microsoft window removes all destination events in mirror mode. Empty results are not treated as errors.
- Uses source identity hashes and private Google event metadata, plus deterministic insert IDs. Interrupted inserts are retried without routinely duplicating meetings. Deleted Google IDs get a persisted replacement generation.
- Updates are applied before removals. Google operations are not transactional: an interrupted write phase may leave partial changes until the next successful run.
- Google-side edits to copied fields are overwritten. Deleted copies reappear while still eligible in Microsoft. Duplicate copies are removed.
- Preserves subjects, descriptions, links, locations, event times, all-day dates, private visibility, and free/busy. Guests and reminders are not copied, so this mirror does not invite attendees or duplicate reminders.
- All-day dates are fetched in their original timezone through Graph, including Windows timezone names and DST. An unsupported timezone or inconsistent event aborts the run instead of risking an incorrect cleanup.
- A script lock prevents overlapping runs. Rate-limit/server failures use bounded retries; a four-minute budget leaves time before Apps Script's runtime limit. Authentication errors do not clear the mirror.

Run `showSyncStatus` to inspect the last successful run, last preview, counts, and error timestamp. Check the Apps Script **Executions** page for errors; Google also provides trigger failure notifications. Logs omit event contents and API response bodies. `authorizeMicrosoft` deliberately logs a short-lived sign-in URL; don't share it.

If authorization fails, check client secret expiration, rotate its value in Script Properties if needed, and run `authorizeMicrosoft` again. To explicitly discard the current user's Microsoft tokens, run `resetMicrosoftAuthorization` before authorizing again. Sign in using the same Google account that owns the trigger. Token revocation and tenant policy can require interactive reauthorization even with automatic refresh. [Microsoft refresh tokens](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens)

Each five-minute schedule produces about 288 runs/day. Consumer Apps Script accounts currently allow 90 minutes/day of trigger runtime (about 18.75 seconds/run on average); Workspace has a larger allowance. All-day events require extra Graph reads, so measure actual execution time and quota usage for your calendar. [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas)

## Development

```bash
npm run check  # source syntax and manifest
npm test       # mocked APIs and reconciliation tests; no credentials required
npm run push   # check, test, then upload source
```

`src/` contains Apps Script globals; Node is used only for local checks, tests, and clasp. API calls use REST through `UrlFetchApp` so retries can inspect status codes and `Retry-After`. The Calendar service declaration enables the underlying Google API; the sync uses its REST endpoints.

Tests cover filtering, recurrence identities, all-day/DST mapping, pagination, reconciliation, retries, interrupted insert recovery, deletion ordering, dry runs, locks, OAuth configuration, and triggers. Live Microsoft/Google authorization and an end-to-end sync must be verified after your credentials are configured.
