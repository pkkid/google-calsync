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
  var start = new Date(o365event.Start)
  var end = new Date(o365event.End)
  var allday = start.getHours() == 0 && start.getMinutes() == 0
  var gcalevent = {}
  if (allday) {
    gcalevent.start = {date:start.toISOString().split('T')[0]}
    gcalevent.end = {date:end.toISOString().split('T')[0]}
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
// Get existing events and eventids from Outlook
function fetchOutlookEvents(url, mindate, maxdate) {
  console.log(`Reading html Outlook calendar ${url}`)
  var caldata = _getAnonymousCalendarSessionData(url)
  return _findItem(url, caldata, mindate, maxdate)
}


// Get Anonymous Calendar Session Data
// Fetch Outlook calendar data from the html url
function _getAnonymousCalendarSessionData(url) {
    var headers = {
      'action': 'GetAnonymousCalendarSessionData',
      'cookie': 'ClientId=7D441659E2204FF2A86A7C6DD0BF04C6; OIDC=1; MSFPC=GUID=fcdcbf52a53a4c209a774984e4babe1a&HASH=fcdc&LV=202512&V=4&LU=1764688053765',
      'origin': 'https://outlook.office365.com',
      'x-owa-canary': 'X-OWA-CANARY_cookie_is_null_or_empty',
      'x-owa-hosted-ux': 'false',
      'x-client-version': '20260724003.16',
      'x-req-source': 'PublishedCalendar',
      'x-owa-sessionid': '3197d5be-39f8-4deb-87c0-cdce7db33b6d',
      'x-owa-correlationid': '4ebfa929-4c6a-d6ea-78d2-f69a098de051',
      'x-owa-actionsource': 'GetAnonymousCalendarSessionData',
      'x-owa-urlpostdata': encodeURIComponent(JSON.stringify({
        "__type": "GetAnonymousCalendarSessionDataJsonRequest:#Exchange",
        "Header": {"__type": "JsonRequestHeaders:#Exchange",
          "RequestServerVersion": "V2018_01_08",
          "TimeZoneContext": {"__type": "TimeZoneContext:#Exchange",
            "TimeZoneDefinition": {"__type": "TimeZoneDefinitionType:#Exchange",
              "Id":"Pacific Standard Time"
      }}}}))
    }
    url = url.replace('/calendar/published/', '/owa/calendar/')
    url = url.replace('/calendar.html', '/service.svc')
    url = url + '?action=GetAnonymousCalendarSessionData&app=PublishedCalendar&n=0'
    console.log(`Fetching Outlook calendar data ${url}`)
    var response = UrlFetchApp.fetch(url, {validateHttpsCertificates:false , headers:headers, muteHttpExceptions:true})
    console.log(response.getContentText())
    return JSON.parse(response.getContentText())
}


// Find Item
// Fetch Outlook calendar events between mindate and maxdate
function _findItem(url, caldata, mindate, maxdate) {
  var dateformat = '%Y-%m-%dT00:00:00.000'
  var headers = {
    'action': 'FindItem',
    'x-owa-urlpostdata': encodeURIComponent(JSON.stringify({
      "__type": "FindItemJsonRequest:#Exchange",
      "Header": {"__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "Exchange2015",
        "TimeZoneContext": {"__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {"__type": "TimeZoneDefinitionType:#Exchange",
          "Id":"Pacific Standard Time"
      }}},
      "Body": {"__type": "FindItemRequest:#Exchange",
        "ParentFolderIds": [{"__type":"FolderId:#Exchange", "Id":"{folderid}"}],
        "ItemShape": {"__type": "ItemResponseShape:#Exchange", "BaseShape": "IdOnly"},
        "Traversal": "Shallow",
        "Paging": {"__type":"CalendarPageView:#Exchange",
          "StartDate": `${mindate.toISOString()}`,
          "EndDate": `${maxdate.toISOString()}`
      }}
    }))
  }
  url = url.replace('/calendar/published/', '/owa/calendar/')
  url = url.replace('/calendar.html', '/service.svc')
  console.log(`Fetching Outlook calendar events ${url}`)
  var response = UrlFetchApp.fetch(url, {method:'POST', validateHttpsCertificates:false , headers:headers, muteHttpExceptions:true})
  var data = JSON.parse(response.getContentText())
  return data.Body.ResponseMessages.Items[0].RootFolder.Items || []
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

















