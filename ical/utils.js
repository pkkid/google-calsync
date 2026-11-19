// Utility Functions for Calendar Sync
// Original Author: Derek Antrican
var RETRY_ERRORS = [
  'service invoked too many times in a short time',
  'rate limit exceeded',
  'internal error'
]


// Create Target Event
// Creates a Google Calendar Event based on the specified ICALEvent.
// Will return null if the event has not changed since the last sync or
// this event should be skipped for another reason.
function createTargetEvent(occurance, gcaltz) {
  var gcalevent = {}
  // Set the start and end times
  if (occurance.startDate.isDate) {
    // All day event
    gcalevent.start = {date: occurance.startDate.toString()}
    gcalevent.end = {date: occurance.endDate.toString()}
  } else {
    // Regular event
    var tzid = occurance.startDate.timezone
    if (tzids.indexOf(tzid) == -1) {
      if (tzid in tzidreplace) { tzid = tzidreplace[tzid] }
      else { tzid = gcaltz }
    }
    gcalevent.start = {dateTime: occurance.startDate.toString(), timeZone: tzid}
    gcalevent.end = {dateTime: occurance.endDate.toString(), timeZone: tzid}
  }
  // Set other basic properties
  gcalevent.summary = occurance.item.summary || occurance.item.description
  // gcalevent.location = occurance.item.location  // OFF FOR NOW!
  gcalevent.description = occurance.item.description
  // Create a uid for the gcalevent and return
  var uid = hashObject(gcalevent)
  gcalevent.extendedProperties = {private: {id:uid}}
  return gcalevent
}


// Events Match
// Returns true if the two gcal events match
function eventsMatch(event1, event2) {
  const dtstr = (str) => str?.match(/(\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d).*?/)[1]
  // Start date and timezone
  if (event1.start.date != event2.start.date) { return false }
  if (dtstr(event1.start.dateTime) != dtstr(event2.start.dateTime)) { return false }
  if (event1.start.timeZone != event2.start.timeZone) { return false }
  // End date and timezone
  if (event1.end.date != event2.end.date) { return false }
  if (dtstr(event1.end.dateTime) != dtstr(event2.end.dateTime)) { return false }
  if (event1.end.timeZone != event2.end.timeZone) { return false }
  // Other fields
  if (event1.summary != event2.summary) { return false }
  if (event1.location != event2.location) { return false }
  if (event1.description != event2.description) { return false }
  return true
}


// Fetch Calendar
// Gets the ressource from the specified ical URL.
function fetchCalendar(url) {
  url = url.replace("webcal://", "https://")
  return retry(function() {
    var response = UrlFetchApp.fetch(url, {'validateHttpsCertificates':false , 'headers':{}})
    if (response.getResponseCode() == 200) {
      var content = RegExp("BEGIN:VCALENDAR.*?END:VCALENDAR", "s").exec(response.getContentText())
      if (content == null) {
        console.log("[ERROR] Incorrect ics/ical URL: " + url)
        return
      }
      return content[0]
    } else {
      throw `Error: Encountered ${response.getReponseCode()} when accessing ${url}`
    }
  })
}


// Fetch Existing Target Events
// Get existing events and eventids from Google Calendar.
function fetchExistingTargetEvents(gcal) {
  var gcalevents = []
  console.log(`Fetching existing events from '${gcal.summary}' calendar`)
  // Fetch the initial page of events
  var response = retry(function() {
    return Calendar.Events.list(gcal.id, {showDeleted:false, maxResults:2500})
  })
  gcalevents = [].concat(gcalevents, response.items)
  // Fetch all remaining pages
  while (typeof response.nextPageToken !== 'undefined') {
    response = retry(function() {
      return Calendar.Events.list(gcal.id, {showDeleted:false, maxResults:2500,
        pageToken:response.nextPageToken})
    })
    gcalevents = [].concat(gcalevents, response.items)
  }
  console.log(`Found ${gcalevents.length} existing events in '${gcal.summary}' calendar`)
  return gcalevents
}

// Hash Object
// hash the passed obj making sure the keys are always sorted
function hashObject(obj) {
  const canonicalJson = JSON.stringify(recSortKeys(obj))
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, canonicalJson)
  return hash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('')
}


// Recursivley Sort Keys
// Make sure all keys in the object are sorted
function recSortKeys(obj) {
  if (obj === null || typeof obj !== 'object') { return obj }
  if (Array.isArray(obj)) { return obj.map(recSortKeys) }
  return Object.keys(obj).sort().reduce((result, key) => {
    result[key] = recSortKeys(obj[key])
    return result
  }, {})
}

// Occurances
// Returns a list of occurance objects
function occurancesInRange(icalevent, mindate, maxdate) {
  var occurances = []
  if (icalevent.isRecurrenceException()) {
    // Returns a single occurance (in a list)
    var start = icalevent.recurrenceId
    if ((start.compare(mindate) >= 0) && (start.compare(maxdate) <= 0)) {
      var occurance = icalevent.getOccurrenceDetails(start)
      occurances.push(occurance)
    }
  } else if (icalevent.isRecurring()) {
    // Returns multiple occurances
    var expand = icalevent.iterator()
    while (start = expand.next()) {
      if ((start.compare(mindate) >= 0) && (start.compare(maxdate) <= 0)) {
        var occurance = icalevent.getOccurrenceDetails(start)
        occurances.push(occurance)
      } else if (start.compare(maxdate) > 0) {
        break
      }
    }
  } else {
    // Regular one time event
    var start = icalevent.startDate
    if ((start.compare(mindate) >= 0) && (start.compare(maxdate) <= 0)) {
      var occurance = {item:icalevent, startDate:start, endDate:icalevent.endDate}
      occurances.push(occurance)
    }
  }
  return occurances
}


// Parse Source Events
// Parse and return ical events
function parseSourceEvents(ical) {
  // Initialize the ICAL utility
  var idata = ICAL.parse(ical)
  var icomponent = new ICAL.Component(idata)
  ICAL.helpers.updateTimezones(icomponent)
  var tzs = icomponent.getAllSubcomponents('vtimezone')
  for (var tz of tzs) { ICAL.TimezoneService.register(tz) }

  // Parse and return all the ical events
  var icalevents = []
  var isubcomponents = icomponent.getAllSubcomponents('vevent')
  isubcomponents.forEach(function(icalcomp) {
    icalcomp.removeProperty('dtstamp')
    var icalevent = new ICAL.Event(icalcomp, {strictExceptions: true})
    icalevents.push(icalevent)
  })
  console.log(`Parsed ${icalevents.length} events from iCalendar`)
  return icalevents
}


// Retry
// Runs the specified function with exponential backoff and returns the result.
function retry(func, maxtries=10) {
  var result
  var tries = 0
  while (tries <= maxtries) {
    tries++
    try {
      result = func()
      return result
    } catch(err) {
      var isrecoverable = RETRY_ERRORS.some(function(e) {
        return err.message.toLowerCase().includes(e)
      })
      if (err.message.includes('is not a function') || !isrecoverable) {
        throw err
      } else if (err.message.includes("Forbidden")) {
        return null
      } else if (tries > maxtries) {
        console.log(`Error, giving up after trying ${maxtries} times [${err}]`)
        return null
      } else {
        console.log( "Error, Retrying... [" + err  +"]")
        Utilities.sleep (Math.pow(2,tries)*100) + (Math.round(Math.random() * 100))
      }
    }
  }
  return null
}


// Setup Target Calendar
// Gets the user's Google Calendar with the specified name.
// A new Calendar will be created if the specified calendar name does not exist.
function setupTargetCalendar(name) {
  console.log(`Setting up Google calendar '${name}'`)
  var targetcal = Calendar.CalendarList.list({showHidden:true}).items.filter(function(cal) {
    return ((cal.summaryOverride || cal.summary) == name) &&
      (cal.accessRole == 'owner' || cal.accessRole == 'writer')
  })[0]
  if (targetcal == null) {
    Logger.log(`Creating calendar ${name}`)
    targetcal = Calendar.newCalendar()
    targetcal.summary = name
    targetcal.description = 'Created by Simple Calendar Sync'
    targetcal.timeZone = Calendar.Settings.get('timezone').value
    targetcal = Calendar.Calendars.insert(targetcal)
  }
  return targetcal
}


// Validate Timer Frequency
// Makes sure we entered a sane number
function validateTimerFreq(freq) {
  if (!freq > 0) {
    console.log("No valid frequency specified. Defaulting to 15 minutes.")
    return 15
  }
  freq = Math.round(freq/5) * 5
  freq = Math.max(freq, 1)
  freq = Math.min(freq, 15)
  return freq
}

