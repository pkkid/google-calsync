function google_(path, token, method, body, deadline, allowed) {
  return request_('https://www.googleapis.com/calendar/v3/' + path, token, method, body, {}, deadline, allowed);
}

function validateTarget_(c, token, deadline) {
  if (!c.targetId || c.targetId === 'primary') throw new Error('Run configureTargetCalendar first; primary is forbidden.');
  const entry = google_('users/me/calendarList/' + encodeURIComponent(c.targetId), token, 'get', undefined, deadline).data;
  if (!entry || entry.primary || entry.id !== c.targetId || entry.summary !== c.targetName ||
      !['owner', 'writer'].includes(entry.accessRole)) {
    throw new Error('Target must be the configured named secondary calendar with write access.');
  }
}

function readDestination_(c, token, deadline) {
  const events = [], pages = new Set();
  let pageToken;
  do {
    const params = { maxResults: 2500, singleEvents: false, showDeleted: false, pageToken: pageToken };
    if (c.mode === 'managed') params.privateExtendedProperty = 'syncOwner=' + c.marker;
    const data = google_('calendars/' + encodeURIComponent(c.targetId) + '/events?' + query_(params),
      token, 'get', undefined, deadline).data;
    if (!data || data.kind !== 'calendar#events' ||
        (data.items !== undefined && !Array.isArray(data.items))) throw new Error('Invalid Google event page.');
    (data.items || []).forEach(function (e) { events.push(e); });
    pageToken = data.nextPageToken;
    if (pageToken && (typeof pageToken !== 'string' || pages.has(pageToken))) throw new Error('Invalid Google pagination.');
    pages.add(pageToken);
  } while (pageToken);
  return events;
}

function eventPath_(c, id) {
  return 'calendars/' + encodeURIComponent(c.targetId) + '/events' + (id ? '/' + encodeURIComponent(id) : '');
}

function updateEvent_(c, token, id, body, deadline) {
  google_(eventPath_(c, id) + '?sendUpdates=none', token, 'put', body, deadline);
}

function createEvent_(c, token, item, deadline, retiredIds) {
  const store = PropertiesService.getScriptProperties();
  const property = 'EVENT_GENERATION_' + digest_(c.targetId + ':' + item.key);
  let generation = Number(store.getProperty(property) || 0);
  if (!Number.isSafeInteger(generation) || generation < 0) throw new Error('Invalid event generation.');
  for (let attempt = 0; attempt < 4; attempt++) {
    const id = digest_(c.targetId + ':' + item.key + ':' + generation);
    if (!retiredIds.has(id)) {
      const result = google_(eventPath_(c) + '?sendUpdates=none', token, 'post',
        Object.assign({ id: id }, item.body), deadline, [409]);
      if (result.status !== 409) return;
      const found = google_(eventPath_(c, id), token, 'get', undefined, deadline, [404, 410]);
      if (found.status === 200 && found.data.status !== 'cancelled') {
        const p = found.data.extendedProperties && found.data.extendedProperties.private || {};
        if (p.sourceKey !== item.key || p.syncOwner !== c.marker) throw new Error('Google event ID collision.');
        if (!(found.data.recurrence || []).length && !(found.data.attendees || []).length) {
          updateEvent_(c, token, id, item.body, deadline);
          return;
        }
      }
    }
    // Deleted Google IDs cannot reliably be reused. Persist a replacement
    // generation BEFORE attempting creation, so interrupted writes are retryable.
    generation++;
    store.setProperty(property, String(generation));
  }
  throw new Error('Event ID recovery exhausted; retry next run.');
}

function applyPlan_(plan, c, token, deadline) {
  const retired = new Set(plan.deletes.map(function (e) { return e.id; }));
  plan.creates.forEach(function (item) { createEvent_(c, token, item, deadline, retired); });
  plan.updates.forEach(function (item) { updateEvent_(c, token, item.id, item.body, deadline); });
  plan.deletes.forEach(function (event) {
    google_(eventPath_(c, event.id) + '?sendUpdates=none', token, 'delete', undefined, deadline, [404, 410]);
  });
}
