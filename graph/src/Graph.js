function graphGet_(url, token, deadline, timezone) {
  if (!/^https:\/\/graph\.microsoft\.com\/v1\.0\//.test(url)) throw new Error('Unexpected Microsoft pagination URL.');
  const zone = timezone || 'UTC';
  if (/["\r\n]/.test(zone)) throw new Error('Invalid Microsoft timezone.');
  return request_(url, token, 'get', undefined, {
    Prefer: 'IdType="ImmutableId", outlook.timezone="' + zone + '"'
  }, deadline).data;
}

function readSource_(c, token, window, deadline) {
  const root = 'https://graph.microsoft.com/v1.0/me/';
  const path = c.sourceId ? 'calendars/' + encodeURIComponent(c.sourceId) : 'calendar';
  const calendar = graphGet_(root + path, token, deadline);
  if (!calendar || !calendar.id || !calendar.owner || !calendar.owner.address) {
    throw new Error('Invalid Microsoft calendar identity.');
  }
  const source = digest_(JSON.stringify([c.tenant, calendar.owner.address.toLowerCase(), calendar.id]));
  let url = root + path + '/calendarView?' + query_({
    startDateTime: window.start, endDateTime: window.end, '$top': 1000
  });
  const events = [], pages = new Set();
  while (url) {
    if (pages.has(url)) throw new Error('Repeated Microsoft pagination URL.');
    pages.add(url);
    const data = graphGet_(url, token, deadline);
    if (!data || !Array.isArray(data.value)) throw new Error('Invalid Microsoft calendar page.');
    data.value.forEach(function (event) { events.push(event); });
    url = data['@odata.nextLink'];
    if (url !== undefined && typeof url !== 'string') throw new Error('Invalid Microsoft nextLink.');
  }
  // Ask Graph to do Windows/IANA timezone conversion for all-day dates. This
  // avoids a stale timezone mapping and UTC date shifts (including DST).
  events.forEach(function (event) {
    if (!eligible_(event) || !event.isAllDay) return;
    const zone = event.originalStartTimeZone;
    if (!zone || !event.id) throw new Error('All-day event lacks source timezone or ID.');
    const local = graphGet_(root + 'events/' + encodeURIComponent(event.id), token, deadline, zone);
    if (!local || local.id !== event.id || local.changeKey !== event.changeKey || !local.isAllDay ||
        !local.start || !local.end) throw new Error('All-day event changed during read; retry next run.');
    event.localDates = { start: local.start.dateTime, end: local.end.dateTime };
  });
  return { source: source, events: events };
}
