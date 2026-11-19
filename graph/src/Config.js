/* All configuration and credentials are entered in Apps Script Project Settings. */
function config_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET'].forEach(function (key) {
    if (!p[key]) throw new Error('Missing Script Property: ' + key);
  });
  if (!/^[a-zA-Z0-9.-]+$/.test(p.MS_TENANT_ID) ||
      /^(common|organizations|consumers)$/i.test(p.MS_TENANT_ID)) {
    throw new Error('MS_TENANT_ID must identify your specific work tenant.');
  }
  const mode = p.SYNC_MODE || 'mirror';
  if (!['mirror', 'managed'].includes(mode)) throw new Error('SYNC_MODE must be mirror or managed.');
  if (p.DRY_RUN && !['true', 'false'].includes(p.DRY_RUN)) throw new Error('DRY_RUN must be true or false.');
  return {
    tenant: p.MS_TENANT_ID, clientId: p.MS_CLIENT_ID, clientSecret: p.MS_CLIENT_SECRET,
    sourceId: p.MS_CALENDAR_ID || '', targetId: p.GOOGLE_CALENDAR_ID || '',
    targetName: p.GOOGLE_CALENDAR_NAME || 'Nasuni', mode: mode,
    dryRun: p.DRY_RUN !== 'false', marker: 'nasuni-graph-sync-v1'
  };
}

function digest_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function deadline_(deadline) {
  if (Date.now() >= deadline) throw new Error('Sync time budget exhausted; retry on the next run.');
}
