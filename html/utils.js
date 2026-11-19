// Utility Functions for Calendar Sync
// Original Author: Derek Antrican
var RETRY_ERRORS = [
  'service invoked too many times in a short time',
  'rate limit exceeded',
  'internal error'
]


// Create Google Calendar Event
// Creates a Google Calendar Event from the o365event info
// gcalevent.location = occurance.item.location  // off for now
function createGoogleCalendarEvent(o365event, gcaltz) {
  // Outlook returns local times with an offset, e.g. 2026-11-12T00:00:00-05:00.
  // All-day events start and end at local midnight.
  var start = new Date(o365event.Start)
  var end = new Date(o365event.End)
  var allday = /T00:00:00/.test(o365event.Start) && /T00:00:00/.test(o365event.End)
  var gcalevent = {}
  if (allday) {
    gcalevent.start = {date:o365event.Start.split('T')[0]}
    gcalevent.end = {date:o365event.End.split('T')[0]}
  } else {
    gcalevent.start = {dateTime:start.toISOString(), timeZone:gcaltz}
    gcalevent.end = {dateTime:end.toISOString(), timeZone:gcaltz}
  }
  gcalevent.summary = o365event.Subject
  gcalevent.description = o365event.Location?.DisplayName
  gcalevent.extendedProperties = {private: {id:o365event.ItemId.Id}}
  return gcalevent
}


// Events Match
// Returns true if the two gcal events match
function eventsMatch(event1, event2) {
  const start1 = event1.start.dateTime || event1.start.date
  const start2 = event2.start.dateTime || event2.start.date
  const end1 = event1.end.dateTime || event1.end.date
  const end2 = event2.end.dateTime || event2.end.date
  // Start date and timezone
  if (start1 != start2) { return false }
  if (end1 != end2) { return false }
  if (event1.summary != event2.summary) { return false }
  if (event1.location != event2.location) { return false }
  if (event1.description != event2.description) { return false }
  return true
}


// Fetch Google Calendar Events
// Get existing events and eventids from Google Calendar.
function fetchGoogleCalendarEvents(gcal) {
  var gcalevents = []
  console.log(`Fetching existing events from '${gcal.summary}' Google calendar`)
  var response = retry(function() {
    return Calendar.Events.list(gcal.id, {showDeleted:false, maxResults:2500})
  })
  gcalevents = [].concat(gcalevents, response.items)
  while (typeof response.nextPageToken !== 'undefined') {
    response = retry(function() {
      return Calendar.Events.list(gcal.id, {showDeleted:false, maxResults:2500,
        pageToken:response.nextPageToken})
    })
    gcalevents = [].concat(gcalevents, response.items)
  }
  console.log(`Found ${gcalevents.length} existing events in '${gcal.summary}' Google calendar`)
  return gcalevents
}


// Fetch Outlook Events
// Get existing events and eventids from Outlook. Mirrors what the OWA published
// calendar page does: fetch the session data to learn the calendar folder id,
// then FindItem against that folder for the date range.
function fetchOutlookEvents(url, mindate, maxdate) {
  var service = _outlookServiceUrl(url)
  console.log('Reading html Outlook calendar')
  var caldata = _getAnonymousCalendarSessionData(service)
  var folderid = caldata.Body.CalendarFolder.FolderId.Id
  return _findItem(service, folderid, mindate, maxdate)
}


// Outlook Service URL
// Published links look like .../owa/calendar/<a>/<b>/calendar.html (which redirects
// to .../calendar/published/<a>/<b>/calendar.html). The OWA web app calls its API at
// https://outlook.office365.com/owa/published/<a>/<b>/service.svc.
function _outlookServiceUrl(url) {
  var match = url.match(/^(https:\/\/[^/]+)\/(?:owa\/calendar|calendar\/published|owa\/published)(\/.+)\/calendar\.html/)
  if (!match) { throw new Error('OUTLOOK_HTML_URL does not look like a published Outlook calendar.html link') }
  return `${match[1]}/owa/published${match[2]}/service.svc`
}


// Outlook Request
// POST a JSON request to the OWA service endpoint and return the parsed response.
function _outlookRequest(service, action, body) {
  var response = UrlFetchApp.fetch(`${service}?action=${action}&app=PublishedCalendar`, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: {action: action},
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  })
  var code = response.getResponseCode()
  if (code != 200) { throw new Error(`Outlook ${action} failed with HTTP ${code}`) }
  return JSON.parse(response.getContentText())
}


// Outlook Request Header
// Standard Exchange JSON header. Times come back in OUTLOOK_TIMEZONE.
function _outlookHeader() {
  return {"__type": "JsonRequestHeaders:#Exchange",
    "RequestServerVersion": "V2018_01_08",
    "TimeZoneContext": {"__type": "TimeZoneContext:#Exchange",
      "TimeZoneDefinition": {"__type": "TimeZoneDefinitionType:#Exchange",
        "Id": OUTLOOK_TIMEZONE
  }}}
}


// Get Anonymous Calendar Session Data
// Returns the published calendar's session data (includes the calendar folder id)
function _getAnonymousCalendarSessionData(service) {
  return _outlookRequest(service, 'GetAnonymousCalendarSessionData', {
    "__type": "GetAnonymousCalendarSessionDataJsonRequest:#Exchange",
    "Header": _outlookHeader(),
  })
}


// Find Item
// Fetch Outlook calendar events between mindate and maxdate
function _findItem(service, folderid, mindate, maxdate) {
  var data = _outlookRequest(service, 'FindItem', {
    "__type": "FindItemJsonRequest:#Exchange",
    "Header": _outlookHeader(),
    "Body": {"__type": "FindItemRequest:#Exchange",
      "ParentFolderIds": [{"__type": "FolderId:#Exchange", "Id": folderid}],
      "ItemShape": {"__type": "ItemResponseShape:#Exchange", "BaseShape": "IdOnly"},
      "Traversal": "Shallow",
      "Paging": {"__type": "CalendarPageView:#Exchange",
        "StartDate": mindate.toISOString(),
        "EndDate": maxdate.toISOString(),
    }}
  })
  var message = data.Body.ResponseMessages.Items[0]
  if (message.ResponseClass != 'Success') { throw new Error(`Outlook FindItem failed: ${message.ResponseCode} ${message.MessageText || ''}`) }
  var items = message.RootFolder.Items || []
  console.log(`Found ${items.length} Outlook events`)
  return items
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
function setupGoogleCalendar(name) {
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

















