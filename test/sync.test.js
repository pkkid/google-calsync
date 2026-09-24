const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');

const window = { start: '2026-09-21T12:00:00.000Z', end: '2026-10-08T12:00:00.000Z' };
const event = (extra = {}) => ({
  id: 'event-1', changeKey: 'v1', isCancelled: false, responseStatus: { response: 'accepted' },
  isAllDay: false, subject: 'Weekly meeting',
  start: { dateTime: '2026-09-24T14:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-09-24T15:00:00.0000000', timeZone: 'UTC' },
  showAs: 'busy', sensitivity: 'normal', ...extra
});
const plain = (value) => JSON.parse(JSON.stringify(value));
function harness(extra = {}) {
  const values = { MS_TENANT_ID: 'tenant-123', MS_CLIENT_ID: 'client', MS_CLIENT_SECRET: 'secret',
    GOOGLE_CALENDAR_ID: 'nasuni@example.com', DRY_RUN: 'false' };
  const store = { getProperties: () => ({ ...values }), getProperty: (k) => values[k] ?? null,
    setProperty: (k, v) => { values[k] = v; }, deleteProperty: (k) => { delete values[k]; } };
  const logs = [], calls = [], sleeps = [];
  const lock = { tryLock: () => true, releaseLock: () => { lock.released = true; } };
  const context = vm.createContext({ console: { log: (...args) => logs.push(args) },
    PropertiesService: { getScriptProperties: () => store, getUserProperties: () => store },
    LockService: { getScriptLock: () => lock, getUserLock: () => ({}) },
    CacheService: { getUserCache: () => ({}) },
    ScriptApp: { getOAuthToken: () => 'google-token' },
    Utilities: { DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (_, text) => [...crypto.createHash('sha256').update(text).digest()],
      sleep: (ms) => sleeps.push(ms) },
    UrlFetchApp: { fetch: (url, options) => { calls.push({ url, options }); throw new Error('Unexpected request'); } },
    ...extra
  });
  fs.readdirSync('src').filter((f) => f.endsWith('.js')).forEach((f) => {
    vm.runInContext(fs.readFileSync('src/' + f, 'utf8'), context, { filename: f });
  });
  return { s: context, values, logs, store, calls, sleeps, lock };
}
function desired(s, events = [event()]) { return s.desiredEvents_(events, 'source', s.config_(), window); }
function response(status, data, headers = {}) {
  return { getResponseCode: () => status, getContentText: () => data === undefined ? '' : JSON.stringify(data),
    getAllHeaders: () => headers };
}

test('status filters exclude only cancelled and own declined events', () => {
  const { s } = harness();
  for (const response of ['accepted', 'tentativelyAccepted', 'notResponded', 'none', 'organizer']) {
    assert.equal(s.eligible_(event({ responseStatus: { response }, attendees: [{ status: { response: 'declined' } }] })), true);
  }
  assert.equal(s.eligible_(event({ isCancelled: true })), false);
  assert.equal(s.eligible_(event({ responseStatus: { response: 'declined' } })), false);
  assert.throws(() => s.eligible_({}), /Invalid/);
});

test('mapping preserves event details, privacy and free time, without invitations', () => {
  const { s } = harness();
  const body = [...desired(s, [event({ sensitivity: 'private', showAs: 'free',
    body: { contentType: 'html', content: '<p>Hello &amp; bye</p><script>bad()</script>' },
    onlineMeeting: { joinUrl: 'https://teams.example/join' }, webLink: 'https://outlook.example/event',
    location: { displayName: 'Room A' } })]).values()][0];
  assert.equal(body.start.dateTime, '2026-09-24T14:00:00.000Z');
  assert.equal(body.visibility, 'private');
  assert.equal(body.transparency, 'transparent');
  assert.match(body.description, /Hello & bye/);
  assert.match(body.description, /Join: https:\/\/teams/);
  assert.doesNotMatch(body.description, /bad/);
  assert.equal(body.location, 'Room A');
  assert.equal(body.attendees, undefined);
  assert.equal(body.reminders.useDefault, false);
});

test('occurrences remain distinct and moved events keep identity', () => {
  const { s } = harness();
  const a = desired(s, [event(), event({ id: 'occurrence-2' })]);
  assert.equal(a.size, 2);
  const moved = desired(s, [event({ subject: 'Renamed', start: { dateTime: '2026-09-24T14:30:00', timeZone: 'UTC' } })]);
  assert.equal([...a.keys()][0], [...moved.keys()][0]);
  assert.throws(() => desired(s, [event(), event()]), /Duplicate/);
});

test('window overlaps are half-open and aging out removes copies', () => {
  const { s } = harness();
  const data = desired(s);
  const existing = [...data.values()].map((body) => ({ ...plain(body), id: 'copy' }));
  const outside = event({ start: { dateTime: '2026-09-21T11:00:00', timeZone: 'UTC' },
    end: { dateTime: '2026-09-21T12:00:00', timeZone: 'UTC' } });
  const empty = desired(s, [outside]);
  assert.equal(empty.size, 0);
  assert.equal(s.reconcile_(empty, existing, s.config_()).deletes.length, 1);
});

test('all-day multi-day dates preserve local dates across DST', () => {
  const { s } = harness();
  const events = [event({ isAllDay: true,
    start: { dateTime: '2026-10-31T04:00:00', timeZone: 'UTC' },
    end: { dateTime: '2026-11-02T05:00:00', timeZone: 'UTC' },
    localDates: { start: '2026-10-31T00:00:00.0000000', end: '2026-11-02T00:00:00.0000000' } })];
  const body = [...s.desiredEvents_(events, 's', s.config_(),
    { start: '2026-10-20T00:00:00.000Z', end: '2026-11-05T00:00:00.000Z' }).values()][0];
  assert.deepEqual(plain(body.start), { date: '2026-10-31' });
  assert.deepEqual(plain(body.end), { date: '2026-11-02' });
  assert.throws(() => s.allDayDate_('2026-10-31T04:00:00'), /midnight/);
  assert.throws(() => s.allDayDate_('2026-02-30T00:00:00'), /Invalid/);
});

test('reconciliation creates once, detects edits and removes duplicates and unrelated events', () => {
  const { s } = harness();
  const want = desired(s), c = s.config_();
  const first = s.reconcile_(want, [], c);
  assert.equal(first.creates.length, 1);
  const copy = { ...plain(first.creates[0].body), id: 'a' };
  assert.equal(s.reconcile_(want, [copy], c).unchanged, 1);
  const plan = s.reconcile_(want, [{ ...copy, summary: 'Local edit' }, { ...copy, id: 'b' }, { id: 'manual' }], c);
  assert.equal(plan.updates.length, 1);
  assert.equal(plan.deletes.length, 2);
  assert.equal(s.reconcile_(want, [{ id: 'manual' }], { ...c, mode: 'managed' }).deletes.length, 0);
});

test('cancelled/declined/deleted source events remove existing copies', () => {
  const { s } = harness();
  const body = plain([...desired(s).values()][0]);
  for (const events of [[], [event({ isCancelled: true })], [event({ responseStatus: { response: 'declined' } })]]) {
    assert.equal(s.reconcile_(desired(s, events), [{ ...body, id: 'copy' }], s.config_()).deletes.length, 1);
  }
});

test('Graph exhausts pagination with immutable-ID headers and does local all-day lookup', () => {
  const { s } = harness();
  const calls = [];
  const e = event({ isAllDay: true, originalStartTimeZone: 'Eastern Standard Time' });
  s.graphGet_ = (url, token, deadline, zone) => {
    calls.push({ url, zone });
    if (url.includes('/events/')) return { ...e,
      start: { dateTime: '2026-09-24T00:00:00' }, end: { dateTime: '2026-09-25T00:00:00' } };
    if (url.endsWith('page2')) return { value: [e] };
    if (url.includes('calendarView')) return { value: [event()], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/page2' };
    return { id: 'calendar', owner: { address: 'work@example.com' } };
  };
  const result = s.readSource_(s.config_(), 'token', window, Infinity);
  assert.equal(result.events.length, 2);
  assert.equal(calls.at(-1).zone, 'Eastern Standard Time');
  assert.equal(result.events[1].localDates.start, '2026-09-24T00:00:00');
});

test('Graph rejects foreign nextLink before sending a bearer token', () => {
  const { s, calls } = harness();
  assert.throws(() => s.graphGet_('https://evil.example/v1.0/', 'secret', Infinity), /Unexpected/);
  assert.equal(calls.length, 0);
});

test('all-day change during read aborts snapshot', () => {
  const { s } = harness();
  const e = event({ isAllDay: true, originalStartTimeZone: 'UTC' });
  s.graphGet_ = (url) => url.includes('calendarView') ? { value: [e] }
    : url.includes('/events/') ? { ...e, changeKey: 'changed' } : { id: 'cal', owner: { address: 'a' } };
  assert.throws(() => s.readSource_(s.config_(), 'token', window, Infinity), /changed during/);
});

test('Google pagination has no time filter so old copies are cleaned up', () => {
  const { s } = harness();
  const urls = [];
  s.google_ = (url) => { urls.push(url); return { data: urls.length === 1
    ? { kind: 'calendar#events', items: [{ id: 'old' }], nextPageToken: 'next' }
    : { kind: 'calendar#events', items: [{ id: 'new' }] } }; };
  assert.equal(s.readDestination_(s.config_(), 'token', Infinity).length, 2);
  assert.match(urls[1], /pageToken=next/);
  assert.doesNotMatch(urls.join(''), /timeMin|timeMax/);
  assert.match(urls[0], /singleEvents=false/);
});

test('HTTP retries rate limits and preserves Retry-After', () => {
  const { s, sleeps } = harness();
  const replies = [response(429, {}, { 'Retry-After': '3' }), response(200, { value: [] })];
  s.UrlFetchApp.fetch = () => replies.shift();
  assert.deepEqual(plain(s.request_('url', 'secret', 'get', undefined, {}, Date.now() + 10000).data), { value: [] });
  assert.deepEqual(sleeps, [3000]);
});

test('HTTP errors redact payloads and do not retry forbidden requests', () => {
  const { s } = harness();
  let count = 0;
  s.UrlFetchApp.fetch = () => { count++; return response(403, { error: 'secret-token' }); };
  assert.throws(() => s.request_('url', 'secret', 'get', undefined, {}, Infinity), /^Error: API HTTP 403\.$/);
  assert.equal(count, 1);
});

test('HTTP retry budget stops long sleeps', () => {
  const { s, sleeps } = harness();
  s.UrlFetchApp.fetch = () => response(429, {}, { 'Retry-After': '300' });
  assert.throws(() => s.request_('url', 'secret', 'get', undefined, {}, Date.now() + 10000), /budget/);
  assert.equal(sleeps.length, 0);
});

test('writes finish before deletions and failed upsert prevents every delete', () => {
  const { s } = harness();
  const order = [];
  const plan = { creates: [{}], updates: [{ id: 'x' }], deletes: [{ id: 'y' }] };
  s.createEvent_ = () => order.push('create');
  s.updateEvent_ = () => { order.push('update'); throw new Error('failed'); };
  s.google_ = () => order.push('delete');
  assert.throws(() => s.applyPlan_(plan, s.config_(), 'token', Infinity), /failed/);
  assert.deepEqual(order, ['create', 'update']);
  s.updateEvent_ = () => order.push('update');
  order.length = 0;
  s.applyPlan_(plan, s.config_(), 'token', Infinity);
  assert.deepEqual(order, ['create', 'update', 'delete']);
});

test('an uncertain successful insert is recovered by deterministic ID without duplication', () => {
  const { s } = harness();
  const item = s.reconcile_(desired(s), [], s.config_()).creates[0];
  const methods = [];
  s.google_ = (url, token, method, body) => {
    methods.push(method);
    if (method === 'post') { assert.match(body.id, /^[0-9a-f]{64}$/); return { status: 409 }; }
    if (method === 'get') return { status: 200, data: { ...plain(item.body), id: 'id' } };
    return { status: 200 };
  };
  s.createEvent_(s.config_(), 'token', item, Infinity, new Set());
  assert.deepEqual(methods, ['post', 'get', 'put']);
});

test('deleted event ID replacement is persisted before insert and reused after interruption', () => {
  const { s, values } = harness();
  const item = s.reconcile_(desired(s), [], s.config_()).creates[0];
  let inserts = 0, replacement;
  s.google_ = (url, token, method, body) => {
    if (method === 'get') return { status: 410 };
    inserts++;
    if (inserts === 1) return { status: 409 };
    assert.equal(Object.entries(values).find(([k]) => k.startsWith('EVENT_GENERATION_'))[1], '1');
    replacement = body.id;
    throw new Error('interrupted');
  };
  assert.throws(() => s.createEvent_(s.config_(), 'token', item, Infinity, new Set()), /interrupted/);
  s.google_ = (url, token, method, body) => { assert.equal(body.id, replacement); return { status: 200 }; };
  s.createEvent_(s.config_(), 'token', item, Infinity, new Set());
});

function runnable() {
  const h = harness(), s = h.s;
  s.validateTarget_ = () => {};
  s.microsoftToken_ = () => 'token';
  s.readSource_ = () => ({ source: 'source', events: [] });
  s.readDestination_ = () => [{ id: 'manual' }];
  return h;
}

test('read failure releases lock, records failure and performs zero writes', () => {
  const { s, values, lock } = runnable();
  let writes = 0;
  s.readSource_ = () => { throw new Error('page two failed'); };
  s.applyPlan_ = () => writes++;
  assert.throws(() => s.syncCalendar(), /page two/);
  assert.equal(writes, 0);
  assert.equal(lock.released, true);
  assert.equal(values.LAST_SUCCESS, undefined);
  assert.ok(values.LAST_ERROR);
});

test('destination page failure also causes zero writes', () => {
  const { s } = runnable();
  s.readDestination_ = () => { throw new Error('Google page failed'); };
  s.applyPlan_ = () => assert.fail('must not write');
  assert.throws(() => s.syncCalendar(), /Google page/);
});

test('preview and default dry-run never mutate calendars or record sync success', () => {
  const { s, values } = runnable();
  s.applyPlan_ = () => assert.fail('must not write');
  const preview = s.previewSync();
  assert.equal(preview.deletes, 1);
  assert.equal(preview.dryRun, true);
  delete values.DRY_RUN;
  assert.equal(s.syncCalendar().dryRun, true);
  assert.equal(values.LAST_SUCCESS, undefined);
});

test('successful empty source deletes destination and records success', () => {
  const { s, values } = runnable();
  let deletes;
  s.applyPlan_ = (plan) => { deletes = plan.deletes.length; };
  s.syncCalendar();
  assert.equal(deletes, 1);
  assert.ok(values.LAST_SUCCESS);
});

test('overlapping runs skip without source reads or writes', () => {
  const { s, lock } = runnable();
  lock.tryLock = () => false;
  s.readSource_ = () => assert.fail('must not read');
  assert.equal(s.syncCalendar().skipped, true);
});

test('primary, renamed and read-only target calendars are rejected', () => {
  const { s } = harness();
  const c = s.config_();
  for (const extra of [{ primary: true }, { summary: 'Other' }, { accessRole: 'reader' }]) {
    s.google_ = () => ({ data: { id: c.targetId, summary: c.targetName, accessRole: 'owner', ...extra } });
    assert.throws(() => s.validateTarget_(c, 'token', Infinity), /Target must/);
  }
  assert.throws(() => s.validateTarget_({ ...c, targetId: 'primary' }, 'token', Infinity), /primary/);
});

test('OAuth uses persistent user storage and refresh lock, and sanitizes auth failures', () => {
  const { s, store } = harness();
  const settings = {};
  const service = new Proxy({}, { get: (_, name) => (value) => { settings[name] = value; return service; } });
  s.OAuth2 = { createService: () => service };
  s.microsoftService_();
  assert.equal(settings.setPropertyStore, store);
  assert.match(settings.setScope, /Calendars.Read offline_access/);
  assert.match(settings.setTokenUrl, /tenant-123\/oauth2\/v2.0\/token/);
  assert.ok(settings.setLock);
  s.OAuth2 = { createService: () => { throw new Error('secret-token'); } };
  assert.throws(() => s.microsoftToken_(), /Microsoft authorization unavailable/);
});

test('trigger installation replaces own sync triggers and requires dry-run disabled', () => {
  const { s, values } = runnable();
  let removed = 0, created = 0;
  s.ScriptApp.getProjectTriggers = () => [{ getHandlerFunction: () => 'syncCalendar' }, { getHandlerFunction: () => 'other' }];
  s.ScriptApp.deleteTrigger = () => removed++;
  s.ScriptApp.newTrigger = (name) => {
    assert.equal(name, 'syncCalendar');
    return { timeBased: () => ({ everyMinutes: (n) => {
      assert.equal(n, 5); return { create: () => created++ };
    } }) };
  };
  s.installTrigger();
  assert.equal(removed, 1);
  assert.equal(created, 1);
  values.DRY_RUN = 'true';
  assert.throws(() => s.installTrigger(), /DRY_RUN=false/);
});

test('complete sync through HTTP creates, removes stale data, and does no writes on repeat', () => {
  const { s, values } = harness();
  const current = event({
    start: { dateTime: new Date(Date.now() + 3600000).toISOString(), timeZone: 'UTC' },
    end: { dateTime: new Date(Date.now() + 7200000).toISOString(), timeZone: 'UTC' }
  });
  const googleEvents = new Map([['old', { id: 'old', summary: 'Old copy' }]]);
  const writes = [];
  s.microsoftToken_ = () => 'microsoft-token';
  s.UrlFetchApp.fetch = (url, options) => {
    if (url.includes('graph.microsoft.com')) {
      assert.equal(options.headers.Authorization, 'Bearer microsoft-token');
      assert.match(options.headers.Prefer, /ImmutableId/);
      return response(200, url.includes('calendarView') ? { value: [current] }
        : { id: 'source-calendar', owner: { address: 'work@example.com' } });
    }
    assert.equal(options.headers.Authorization, 'Bearer google-token');
    if (url.includes('/calendarList/')) return response(200, {
      id: values.GOOGLE_CALENDAR_ID, summary: 'Nasuni', accessRole: 'owner'
    });
    if (options.method === 'get') return response(200, { kind: 'calendar#events', items: [...googleEvents.values()] });
    writes.push(options.method);
    if (options.method === 'post') {
      const body = JSON.parse(options.payload);
      googleEvents.set(body.id, { ...body, reminders: { useDefault: false, overrides: [] } });
      return response(200, body);
    }
    assert.equal(options.method, 'delete');
    googleEvents.delete('old');
    return response(204);
  };
  const first = s.syncCalendar();
  assert.equal(first.creates, 1);
  assert.equal(first.deletes, 1);
  assert.deepEqual(writes, ['post', 'delete']);
  const again = s.syncCalendar();
  assert.equal(again.unchanged, 1);
  assert.deepEqual(writes, ['post', 'delete']);
  assert.ok(values.LAST_SUCCESS);
});

test('later Graph HTTP page failure aborts complete sync before all Google writes', () => {
  const { s, values } = harness();
  s.microsoftToken_ = () => 'token';
  s.UrlFetchApp.fetch = (url, options) => {
    assert.equal(options.method, 'get');
    if (url.includes('/calendarList/')) return response(200, {
      id: values.GOOGLE_CALENDAR_ID, summary: 'Nasuni', accessRole: 'owner'
    });
    if (url.endsWith('page2')) return response(403, { error: 'denied' });
    if (url.includes('calendarView')) return response(200, {
      value: [event()], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/page2'
    });
    return response(200, { id: 'source-calendar', owner: { address: 'work@example.com' } });
  };
  assert.throws(() => s.syncCalendar(), /403/);
  assert.equal(values.LAST_SUCCESS, undefined);
});

test('malformed successful Google response does not count as an empty calendar', () => {
  const { s } = harness();
  s.google_ = () => ({ data: {} });
  assert.throws(() => s.readDestination_(s.config_(), 'token', Infinity), /Invalid Google/);
});

test('recurrence or guests added to a copy cause replacement before deletion', () => {
  const { s } = harness();
  const want = desired(s), body = plain([...want.values()][0]);
  for (const extra of [{ recurrence: ['RRULE:FREQ=DAILY'] }, { attendees: [{ email: 'guest@example.com' }] }]) {
    const plan = s.reconcile_(want, [{ ...body, id: 'modified', ...extra }], s.config_());
    assert.equal(plan.creates.length, 1);
    assert.equal(plan.deletes[0].id, 'modified');
  }
});
