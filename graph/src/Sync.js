function previewSync() { return runSync_(true); }
function syncCalendar() { return runSync_(false); }

function runSync_(preview) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1)) { console.log('Sync skipped: another run is active.'); return { skipped: true }; }
  const started = Date.now(), deadline = started + 240000;
  const store = PropertiesService.getScriptProperties();
  try {
    const c = config_(), dryRun = preview || c.dryRun;
    const window = {
      start: new Date(started - 3 * 86400000).toISOString(),
      end: new Date(started + 14 * 86400000).toISOString()
    };
    const googleToken = ScriptApp.getOAuthToken();
    validateTarget_(c, googleToken, deadline);
    const source = readSource_(c, microsoftToken_(), window, deadline);
    const desired = desiredEvents_(source.events, source.source, c, window);
    const existing = readDestination_(c, googleToken, deadline);
    const plan = reconcile_(desired, existing, c);
    const summary = { dryRun: dryRun, mode: c.mode, window: window,
      sourceEvents: source.events.length, eligibleEvents: desired.size,
      creates: plan.creates.length, updates: plan.updates.length,
      deletes: plan.deletes.length, unchanged: plan.unchanged };
    if (!dryRun) {
      applyPlan_(plan, c, googleToken, deadline);
      store.setProperty('LAST_SUCCESS', new Date().toISOString());
      store.deleteProperty('LAST_ERROR');
    }
    summary.durationMs = Date.now() - started;
    store.setProperty(dryRun ? 'LAST_PREVIEW' : 'LAST_RUN', JSON.stringify(summary));
    console.log(JSON.stringify(summary));
    return summary;
  } catch (error) {
    // All our network/OAuth errors are sanitized before reaching this boundary.
    // Do not store arbitrary exception messages from external Apps Script services.
    store.setProperty('LAST_ERROR', JSON.stringify({ at: new Date().toISOString(), message: 'Sync failed. Check Executions.' }));
    throw error;
  } finally { lock.releaseLock(); }
}

function showSyncStatus() {
  const p = PropertiesService.getScriptProperties();
  const status = {};
  ['LAST_SUCCESS', 'LAST_RUN', 'LAST_PREVIEW', 'LAST_ERROR'].forEach(function (key) { status[key] = p.getProperty(key); });
  console.log(JSON.stringify(status));
  return status;
}
