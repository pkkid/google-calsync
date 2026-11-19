# google-calsync

Google Apps Script projects that copy an Outlook (Office 365) work calendar into the
**Nasuni** Google Calendar. Each folder is a separate Apps Script project, pushed with
[clasp](https://github.com/google/clasp).

| Folder | Source | Apps Script project | Status |
| --- | --- | --- | --- |
| `html/` | OWA published calendar (HTML view, `service.svc` API) | Simple Calendar Sync (`1Bxn...`) | Preferred |
| `ical/` | OWA published calendar `.ics` feed | Simple Calendar Sync (`1WNN...`) | Fallback, unreliable |
| `graph/` | Microsoft Graph API (needs Entra app registration) | not deployed | Experimental, see `graph/README.md` |

All of them write to the same Google calendar, so **only one project should have its
timer trigger installed** at a time (Apps Script editor, Triggers tab).

## Setup

```bash
npm ci
npm run login        # once; stores credentials in ~/.clasprc.json
```

The Apps Script API must be on at https://script.google.com/home/usersettings.

## Secrets

This repo is public. The published-calendar URL grants read access to the work calendar,
so it lives in each project's **Project Settings > Script Properties**, never in code:

| Project | Property | Value |
| --- | --- | --- |
| html | `OUTLOOK_HTML_URL` | `https://outlook.office365.com/owa/calendar/.../calendar.html` |
| ical | `OUTLOOK_ICAL_URL` | `https://outlook.office365.com/owa/calendar/.../calendar.ics` |

`npm run check` fails if a calendar URL appears in source.

## Workflow

```bash
npm run status:html   # which files would upload
npm run push:html     # check, then upload html/ to its Apps Script project
npm run open:html     # open it in the editor (run sync / install there)
npm run pull:html     # pull edits made in the web editor back down
```

Same commands exist with `:ical`. Pushing replaces the server code; triggers and Script
Properties are kept. Run `install` once from the editor to add the timer trigger.
