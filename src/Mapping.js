function eligible_(event) {
  if (!event || typeof event.isCancelled !== 'boolean' || !event.responseStatus ||
      !['none', 'organizer', 'tentativelyAccepted', 'accepted', 'declined', 'notResponded']
        .includes(event.responseStatus.response)) throw new Error('Invalid Microsoft event status.');
  return !event.isCancelled && event.responseStatus.response !== 'declined';
}

function instant_(value) {
  if (!value || typeof value.dateTime !== 'string' || value.timeZone !== 'UTC') {
    throw new Error('Expected UTC Microsoft date-time.');
  }
  const text = value.dateTime.replace(/(\.\d{3})\d+/, '$1');
  const date = new Date(/[zZ]$|[+-]\d\d:\d\d$/.test(text) ? text : text + 'Z');
  if (!Number.isFinite(date.getTime())) throw new Error('Invalid event date-time.');
  return date.toISOString();
}

function allDayDate_(text) {
  if (typeof text !== 'string' || !/^\d{4}-\d\d-\d\dT00:00:00(?:\.0+)?$/.test(text)) {
    throw new Error('All-day event is not midnight in its original timezone.');
  }
  const date = text.slice(0, 10);
  if (new Date(date + 'T00:00:00Z').toISOString().slice(0, 10) !== date) throw new Error('Invalid all-day date.');
  return date;
}

function plainBody_(body) {
  if (!body || !body.content) return '';
  if (String(body.contentType).toLowerCase() !== 'html') return body.content;
  return body.content.replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, '')
    .replace(/<a\b[^>]*href=["'](https?:\/\/[^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi, '$2 ($1)')
    .replace(/<\s*br\s*\/?\s*>|<\/(?:p|div|li|tr|h[1-6])>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, function (_, entity) {
      const key = entity.toLowerCase();
      const known = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
      if (known[key]) return known[key];
      const n = key.startsWith('#x') ? parseInt(key.slice(2), 16) : Number(key.slice(1));
      return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : '';
    }).replace(/\n{3,}/g, '\n\n').trim();
}

function desiredEvents_(events, source, c, window) {
  const result = new Map();
  events.forEach(function (event) {
    if (!eligible_(event)) return;
    if (!event.id || typeof event.isAllDay !== 'boolean') throw new Error('Invalid source event identity.');
    const start = instant_(event.start), end = instant_(event.end);
    if (end <= start) throw new Error('Source event has an invalid time range.');
    if (end <= window.start || start >= window.end) return;
    const key = digest_(source + ':' + event.id);
    if (result.has(key)) throw new Error('Duplicate source occurrence; retry next run.');
    const description = [plainBody_(event.body),
      event.onlineMeeting && event.onlineMeeting.joinUrl ? 'Join: ' + event.onlineMeeting.joinUrl : '',
      event.webLink ? 'Outlook: ' + event.webLink : ''].filter(Boolean).join('\n\n');
    const body = {
      summary: event.subject || '(No title)', description: description,
      location: event.location && event.location.displayName || '',
      start: { dateTime: start }, end: { dateTime: end },
      transparency: event.showAs === 'free' ? 'transparent' : 'opaque',
      visibility: ['private', 'confidential'].includes(event.sensitivity) ? 'private' : 'default',
      status: 'confirmed',
      reminders: { useDefault: false },
      extendedProperties: { private: { syncOwner: c.marker, source: source, sourceKey: key } }
    };
    if (event.isAllDay) {
      if (!event.localDates) throw new Error('Missing all-day local dates.');
      body.start = { date: allDayDate_(event.localDates.start) };
      body.end = { date: allDayDate_(event.localDates.end) };
      if (body.end.date <= body.start.date) throw new Error('Invalid all-day range.');
    }
    result.set(key, body);
  });
  return result;
}

function comparable_(event) {
  function time(value) {
    if (!value) return null;
    if (value.date) return { date: value.date };
    const date = new Date(value.dateTime);
    return { dateTime: Number.isFinite(date.getTime()) ? date.toISOString() : 'invalid' };
  }
  const p = event.extendedProperties && event.extendedProperties.private || {};
  return JSON.stringify({
    summary: event.summary || '', description: event.description || '', location: event.location || '',
    start: time(event.start), end: time(event.end),
    transparency: event.transparency || 'opaque', visibility: event.visibility || 'default',
    reminders: {
      useDefault: event.reminders ? event.reminders.useDefault === true : true,
      overrides: (event.reminders && event.reminders.overrides || []).map(function (r) {
        return { method: r.method, minutes: r.minutes };
      }).sort(function (a, b) { return JSON.stringify(a).localeCompare(JSON.stringify(b)); })
    },
    syncOwner: p.syncOwner, source: p.source, sourceKey: p.sourceKey,
    recurrence: event.recurrence || [], attendees: event.attendees || [],
    status: event.status || 'confirmed'
  });
}

function reconcile_(desired, existing, c) {
  const groups = new Map(), deletes = [], creates = [], updates = [];
  let unchanged = 0;
  existing.forEach(function (event) {
    if (!event.id) throw new Error('Google event missing ID.');
    if (event.status === 'cancelled') return;
    const p = event.extendedProperties && event.extendedProperties.private || {};
    if (p.syncOwner !== c.marker) {
      if (c.mode === 'mirror') deletes.push(event);
      return;
    }
    if (!desired.has(p.sourceKey)) { deletes.push(event); return; }
    if (!groups.has(p.sourceKey)) groups.set(p.sourceKey, []);
    groups.get(p.sourceKey).push(event);
  });
  desired.forEach(function (body, key) {
    const matches = (groups.get(key) || []).sort(function (a, b) { return a.id.localeCompare(b.id); });
    // Recurrence or guests indicate a structurally modified copy: replace it.
    const valid = matches.filter(function (e) { return !(e.recurrence || []).length && !(e.attendees || []).length; });
    const keep = valid[0];
    matches.forEach(function (e) { if (e !== keep) deletes.push(e); });
    if (!keep) creates.push({ key: key, body: body });
    else if (comparable_(keep) !== comparable_(body)) updates.push({ id: keep.id, body: body });
    else unchanged++;
  });
  return { creates: creates, updates: updates, deletes: deletes, unchanged: unchanged };
}
