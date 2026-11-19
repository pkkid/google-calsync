// Sync Office365 HTML Calendar to Google Calendar
// -------
// SYNC_CONFIGS: List of {ical, gcal} pairs.
//   * html should be a link to a shared icalendar to sync from.
//   * gcal is the name of the calendar to sync to.
// TIMER_FREQ: When running install, the timer will be setup with this frequency.
// DAYS_TO_SYNC_FROM_TODAY: How many days to sync look back and forward.
// IGNORE_EVENT_PATTERNS: List of event summary regex patterns to ignore during sync.
// The Outlook URL is read from Script Properties (Project Settings) so it
// never lives in source control. Set OUTLOOK_HTML_URL there before running.
function getSyncConfigs() {
  return [{
    html: getRequiredProperty('OUTLOOK_HTML_URL'),
    gcal: 'Nasuni'
  }]
}

function getRequiredProperty(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key)
  if (!value) { throw new Error(`Missing Script Property '${key}'. Add it under Project Settings > Script Properties.`) }
  return value
}
const TIMER_FREQ = 15
const OUTLOOK_TIMEZONE = 'Eastern Standard Time' // Windows tz id; should match appsscript.json timeZone
const DAYS_TO_SYNC_FROM_TODAY = [-7, 60]
const IGNORE_EVENT_PATTERNS = [/^Canceled:/, /^Declined:/, /^Lunch$/]


// Sync
// Run the Sync Script!
function sync() {
  const dryrun = false // Set true to disable modifications (for debugging)
  const mindate = new Date(Date.now() + DAYS_TO_SYNC_FROM_TODAY[0]*24*60*60*1000)
  const maxdate = new Date(Date.now() + DAYS_TO_SYNC_FROM_TODAY[1]*24*60*60*1000)

  for (var config of getSyncConfigs()) {
    console.log(`Syncing Google calendar '${config.gcal}' from Outlook`)
    var o365events = fetchOutlookEvents(config.html, mindate, maxdate)
    var gcal = setupGoogleCalendar(config.gcal)
    var gcaltz = retry(function() { return Calendar.Settings.get('timezone').value })
    var gcalexisting = fetchGoogleCalendarEvents(gcal).reduce((obj, event) => {
      obj[event.extendedProperties.private['id']] = event
      return obj
    }, {})

    // Insert or update events in Google Calendar
    const alluids = new Set()
    o365events.forEach(function(o365event) {
      // Check we are ignoring this event
      if (IGNORE_EVENT_PATTERNS.some(pattern => pattern.test(o365event.Subject))) {
        console.log(`Ignoring event ${o365event.Subject} ${o365event.Start}`);
        return
      }
      // Create the gcal event
      var gcalevent = createGoogleCalendarEvent(o365event, gcaltz)
      var uid = gcalevent.extendedProperties.private['id']
      if (!alluids.has(uid)) {
        alluids.add(uid)
        if (uid in gcalexisting) {
          if (eventsMatch(gcalevent, gcalexisting[uid])) {
            console.log(`Skipping exact match update ${gcalevent.summary} ${gcalevent.start.dateTime || gcalevent.start.date}`)
            return
          }
          console.log(`Updating existing event ${gcalevent.summary} ${gcalevent.start.dateTime || gcalevent.start.date}`)
          if (dryrun) { return }
          retry(function() { return Calendar.Events.update(gcalevent, gcal.id, gcalexisting[uid].id) })
        } else {
          console.log(`Adding new event ${gcalevent.summary} ${gcalevent.start.dateTime || gcalevent.start.date}`)
          if (dryrun) { return }
          retry(function() { return Calendar.Events.insert(gcalevent, gcal.id) })
        }
      }
    })

    // Delete all events we did not add or update
    for (const [uid, gcalevent] of Object.entries(gcalexisting)) {
      if (!alluids.has(uid)) {
        console.log(`Deleting old event ${gcalevent.summary} ${gcalevent.start.dateTime || gcalevent.start.date} - ${uid}`)
        if (dryrun) { continue }
        retry(function() { Calendar.Events.remove(gcal.id, gcalevent.id) })
      }
    }

    // All Done!
    console.log('Done.')
  }
}

// Install
// Sets up new timer trigger
function install() {
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(TIMER_FREQ).create()
  console.log('Installed time based trigger. View it on the Triggers tab.')
}















