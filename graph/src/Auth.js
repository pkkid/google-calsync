function microsoftService_() {
  const c = config_();
  const base = 'https://login.microsoftonline.com/' + c.tenant + '/oauth2/v2.0/';
  return OAuth2.createService('microsoft-' + digest_(c.tenant + ':' + c.clientId).slice(0, 20))
    .setAuthorizationBaseUrl(base + 'authorize').setTokenUrl(base + 'token')
    .setClientId(c.clientId).setClientSecret(c.clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache()).setLock(LockService.getUserLock())
    .setScope('https://graph.microsoft.com/Calendars.Read offline_access')
    .setParam('prompt', 'select_account');
}

function authorizeMicrosoft() {
  const service = microsoftService_();
  console.log('Register this Web redirect URI: ' + service.getRedirectUri());
  console.log('Open this temporary authorization URL: ' + service.getAuthorizationUrl());
}

function authCallback(request) {
  let ok = false;
  try { ok = microsoftService_().handleCallback(request); } catch (_) { /* Never expose token responses. */ }
  return HtmlService.createHtmlOutput(ok
    ? 'Microsoft authorization complete. You can close this tab and run previewSync in Apps Script.'
    : 'Authorization failed. Check tenant consent and credentials, then run authorizeMicrosoft again.');
}

function microsoftToken_() {
  try {
    const service = microsoftService_();
    if (service.hasAccess()) return service.getAccessToken();
  } catch (_) { /* OAuth errors can contain credentials; replace them. */ }
  throw new Error('Microsoft authorization unavailable. Check client secret expiry and run authorizeMicrosoft.');
}

function resetMicrosoftAuthorization() {
  microsoftService_().reset();
}
