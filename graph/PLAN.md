# Office 365 → Nasuni calendar sync

Original design, retained for context. Implemented September 24, 2026; see README.md for the actual setup and behavior. No live calendar changes have been made.

Implementation decisions: dedicated `mirror` mode is the default, with optional `managed` mode and dry-run enabled until configured. Both calendars use REST for explicit retry/status handling; the Advanced Calendar declaration enables the Google API. All-day timezone conversion is delegated to Graph. Event metadata stores a digest of source identity rather than raw Microsoft IDs. App registration, credentials, and live validation remain user setup steps.

## Intended behavior

- Microsoft is the source of truth; sync only toward Google Calendar `Nasuni`.
- Fetch events overlapping `[now − 3 days, now + 14 days)`, using one captured timestamp per run. Proposed interpretation: rolling 24-hour days, not whole local calendar dates.
- Exclude `isCancelled === true` and `responseStatus.response === "declined"`. This checks the owner's response, not whether another attendee declined. Keep accepted, tentative, unanswered, and organizer events. Use structured fields, not title searches. See [Graph event fields](https://learn.microsoft.com/en-us/graph/api/resources/event?view=graph-rest-1.0) and [response statuses](https://learn.microsoft.com/en-us/graph/api/resources/responsestatus?view=graph-rest-1.0).
- Create missing events, update changed events, and remove destination events absent from the latest successful source snapshot—including previously copied events that have aged out of the window.
- Proposed deletion scope: `Nasuni` is a dedicated exact mirror, so manually added and preexisting unrelated events on it are also removed. This interpretation is pending the user's answer. Alternative: restrict deletion to events carrying this script's ownership marker.
- Default source: the work account's primary calendar. Allow an explicit source calendar ID.
- Use `America/New_York` as the proposed display/project timezone; preserve actual event instants and all-day dates.

## Runtime and sync algorithm

Use a standalone Apps Script project, plain JavaScript, `UrlFetchApp` for Microsoft Graph, and the Advanced Google Calendar service for Google event operations.

1. Acquire a script lock; skip if another sync is active. Validate configuration, authentication, and the pinned target calendar ID. Reject the Google primary calendar as a destination. Resolve `Nasuni` by name during setup only; if multiple calendars match, require an explicit ID.
2. Capture the window once. Read every page from `GET /v1.0/me/calendar/calendarView?startDateTime=...&endDateTime=...`, or the configured calendar's equivalent. Follow every `@odata.nextLink`, retaining required headers. Request immutable Graph IDs and UTC responses. `calendarView` expands recurring occurrences and exceptions. See [calendarView](https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0).
3. Validate the entire response and normalize eligible events into a desired-state map. A page error, invalid event, permission error, or authentication failure aborts reconciliation. Never interpret a failed request as an empty calendar. A valid, completely fetched empty result is a legitimate empty mirror.
4. Read every page of destination events before mutations. In exact-mirror mode, list all live events on the pinned calendar without date limits, with recurring series unexpanded, so old events and unrelated recurring series can be removed. In managed-only mode, filter by the script marker, still without date limits. Google supports private-property filters and paginated event listing. See [Events.list](https://developers.google.com/workspace/calendar/api/v3/reference/events/list).
5. Compute creates, updates, unchanged items, duplicates, and deletions. `previewSync()` returns this same plan without writes. Compare actual normalized destination fields with desired fields so Google-side edits are corrected too.
6. Apply creates and updates first. On an upsert failure, stop before deletion. Then remove stale events and duplicates from the captured destination inventory. Never clear and rebuild the whole calendar.
7. Record success only after all required writes finish. A partial write failure is recoverable on the next run; this is convergent reconciliation, not an atomic transaction. Always release the lock.

Retry transient errors with bounded backoff and honor Microsoft's `Retry-After`. Abort before Apps Script's execution deadline. If retries exhaust the budget, keep existing destination data and let the next scheduled run retry. Emit redacted operation counts, duration, last success, and actionable errors; do not log tokens or full calendar payloads.

## Event identity and contents

Store private extended properties on each Google event: an ownership marker, source mailbox/calendar identity, and immutable Graph event/occurrence ID. Use a digest of that composite identity as the compact lookup key. Never match by title or current start time: both can change. Copy each recurrence occurrence as a standalone Google event; no Google recurrence rules are needed.

For retry-safe creates, use a stable, valid Google event ID derived from the composite source key. On a conflict or uncertain insert result, retrieve and reconcile it. Include an explicit recovery path for a previously deleted destination ID: allocate and persist a replacement ID before retrying creation. Google constrains custom IDs to base32hex-compatible characters; hex digests fit. See [Events.insert](https://developers.google.com/workspace/calendar/api/v3/reference/events/insert).

Proposed field mapping:

| Microsoft | Google |
| --- | --- |
| Subject | Summary |
| Start/end | Date-times with explicit offsets |
| All-day start/end | Date-only start and exclusive end |
| Location | Location text |
| Body and Teams join URL | Description, with an Outlook link |
| `showAs` | Free → transparent; other statuses → opaque |
| Private/confidential sensitivity | Private visibility |

Normalize HTML descriptions consistently. For all-day events, recover dates in the source event's original timezone rather than truncating UTC dates; map Windows timezone names where necessary. Preserve multi-day exclusive end dates and test daylight-saving transitions. Omit attendee lists to avoid creating invitations; disable copied reminders by default to avoid duplicate notifications.

## Microsoft authentication and secrets

Preferred design for one mailbox: a single-tenant Entra app with delegated `Calendars.Read` and `offline_access`, using authorization-code OAuth. `Calendars.Read` supports copying event descriptions; no Microsoft write permission is needed. Register the library's callback as a **Web** redirect URI, and use tenant-specific authorization/token endpoints. Initial sign-in and consent happen interactively. See [Microsoft authorization-code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow).

Use a pinned version of Google's Apps Script OAuth2 library for the callback, token persistence, and refresh. Its callback URI is `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback`; no public web-app deployment is needed. Give token storage, caching, and token-refresh locking the same user scope. Keep the overall sync lock separate. See [Apps Script OAuth2 library](https://github.com/googleworkspace/apps-script-oauth2).

| Value | Storage |
| --- | --- |
| Tenant ID, client ID, target/source calendar IDs, settings | Apps Script Script Properties |
| Microsoft client secret and its expiry date | Apps Script Script Properties, entered in Project Settings |
| Access/refresh tokens and expiry metadata | Trigger owner's User Properties, maintained by OAuth library |
| Local `clasp` login credentials | Local clasp credential file, outside version control |
| Local script binding | Ignored `.clasp.json`; commit a placeholder example only |

Properties persist independently of source-code uploads. Script Properties are shared configuration, not an isolated secret vault: anyone able to edit this project must be trusted with its credentials. Keep the project private. Google Cloud Secret Manager is an optional later upgrade if centralized secret access and auditing become necessary. See [Properties Service](https://developers.google.com/apps-script/guides/properties).

Save replacement refresh tokens when returned. Automatic refresh supports unattended operation, but expiration, revocation, and tenant policy can require reauthorization. On auth failure, fail visibly and leave Google events intact. Rotate the client secret before expiry. The setup guide must explain both procedures. See [Microsoft refresh tokens](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens).

IT may need to permit app registration or grant consent. This is the main external prerequisite. App-only mailbox access is a possible IT-managed alternative, requiring a separately scoped design; do not assume it is available.

## Schedule and code uploads

`installTrigger()` should idempotently create one trigger for the owner:

```javascript
ScriptApp.newTrigger('syncCalendar').timeBased().everyMinutes(5).create();
```

Install it using the same Google account that completed Microsoft authorization and can write to `Nasuni`. This schedules nominal five-minute intervals; it is not a strict timing guarantee. Provide `removeTrigger()` as well. See [clock triggers](https://developers.google.com/apps-script/reference/script/clock-trigger-builder) and [trigger execution identity](https://developers.google.com/apps-script/guides/triggers/installable).

At 288 runs/day, consumer accounts' current 90-minute daily trigger allowance implies an average below 18.75 seconds/run, with headroom needed for other scripts. Workspace currently allows six hours/day; individual executions are limited to six minutes. Avoid unnecessary event writes and measure actual runtime before leaving the trigger enabled. See [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas).

Use Google's `clasp`, pinned as a development dependency with a lockfile. Planned npm commands: `npm run push` → `clasp push`, `npm run open` → `clasp open-script`, plus local lint/test commands. Enable the Apps Script API, run `clasp login`, and bind/create a standalone project once. Routine uploads need no new web-app deployment; triggers use the project's saved code. See [clasp guide](https://developers.google.com/apps-script/guides/clasp).

Keep runtime files under `src/`, use that as clasp's root, and allowlist uploadable source/manifest files. Add `.gitignore` entries for `.clasp.json`, clasp credentials, `.env*` except examples, local secret files, and `node_modules/`. `.claspignore` independently prevents accidental secret uploads. Do not create a source file containing real credentials, even temporarily for setup.

## Implementation and validation sequence

1. Scaffold `src/{Config,Auth,Graph,Mapping,Calendar,Sync,Setup}.js`, `src/appsscript.json`, package scripts, ignore files, tests, and a setup README. Enable Advanced Calendar and explicit Google scopes needed for calendar operations, external requests, and trigger management.
2. Implement and test pure filtering, date conversion, identity, and reconciliation logic with synthetic fixtures.
3. Implement authentication and read-only source/destination access, then the preview function.
4. Implement writes, recovery, deletion ordering, scheduling, and redacted status reporting.
5. Configure Entra and Apps Script properties, authorize both accounts, and inspect a preview. Validate writes against a separate test calendar, then switch to the pinned `Nasuni` ID and install the trigger.

Acceptance cases: create/update/no-op repeat; cancel/decline/delete; recurring exceptions and a single canceled occurrence; moved events; aging out; all-day and multi-day dates across DST; Google-side edits/deletions; duplicate recovery; multiple API pages; empty successful source; failed later page causing zero writes; failed upsert causing zero deletions; interrupted create/delete recovery; token refresh/reauthorization; and overlapping triggers.

The implementation defaults to a dedicated mirror, as described above. Entra app registration/IT approval and live setup are still required before deployment can run successfully.
