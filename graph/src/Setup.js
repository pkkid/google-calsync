function configureTargetCalendar() {
  const c = config_(), token = ScriptApp.getOAuthToken(), deadline = Date.now() + 60000;
  if (c.targetId) {
    validateTarget_(c, token, deadline);
    console.log('Configured target calendar validated.');
    return;
  }
  const matches = [], seen = new Set();
  let pageToken;
  do {
    const data = google_('users/me/calendarList?' + query_({ maxResults: 250, pageToken: pageToken }),
      token, 'get', undefined, deadline).data;
    (data.items || []).forEach(function (entry) {
      if (entry.summary === c.targetName && !entry.primary && ['owner', 'writer'].includes(entry.accessRole)) matches.push(entry);
    });
    pageToken = data.nextPageToken;
    if (pageToken && seen.has(pageToken)) throw new Error('Repeated calendar-list page.');
    seen.add(pageToken);
  } while (pageToken);
  if (matches.length !== 1) throw new Error('Expected one writable named secondary calendar. Create it or set GOOGLE_CALENDAR_ID explicitly.');
  PropertiesService.getScriptProperties().setProperty('GOOGLE_CALENDAR_ID', matches[0].id);
  console.log('Target calendar ID saved. Run previewSync.');
}

function installTrigger() {
  if (config_().dryRun) throw new Error('Set DRY_RUN=false after reviewing previewSync, then install the trigger.');
  // Verify complete read access before enabling unattended writes.
  const preview = previewSync();
  if (preview.skipped) throw new Error('A sync is active; retry trigger installation.');
  removeTrigger();
  ScriptApp.newTrigger('syncCalendar').timeBased().everyMinutes(5).create();
  console.log('Five-minute sync trigger installed for this Google account.');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'syncCalendar') ScriptApp.deleteTrigger(trigger);
  });
}
