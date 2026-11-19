// Both APIs use REST so status codes and Retry-After remain available to us.
function request_(url, token, method, body, headers, deadline, allowed) {
  for (let attempt = 0; attempt < 4; attempt++) {
    deadline_(deadline);
    let response;
    try {
      response = UrlFetchApp.fetch(url, {
        method: method || 'get', muteHttpExceptions: true, followRedirects: false,
        headers: Object.assign({ Authorization: 'Bearer ' + token }, headers || {}),
        contentType: 'application/json',
        ...(body === undefined ? {} : { payload: JSON.stringify(body) })
      });
    } catch (_) {
      if (attempt === 3) throw new Error('API network request failed.');
      pauseRetry_(1000 * Math.pow(2, attempt), deadline);
      continue;
    }
    const status = response.getResponseCode();
    if ((allowed || []).includes(status)) return { status: status, data: null };
    if (status >= 200 && status < 300) {
      const text = response.getContentText();
      try { return { status: status, data: text ? JSON.parse(text) : null }; }
      catch (_) { throw new Error('API returned invalid JSON.'); }
    }
    let rateLimit = false;
    if (status === 403) {
      try {
        const data = JSON.parse(response.getContentText());
        rateLimit = (data.error.errors || []).some(function (e) {
          return ['rateLimitExceeded', 'userRateLimitExceeded'].includes(e.reason);
        });
      } catch (_) { /* Not a Google rate-limit response. */ }
    }
    if ((status === 429 || status >= 500 || rateLimit) && attempt < 3) {
      const h = response.getAllHeaders();
      const key = Object.keys(h).find(function (k) { return k.toLowerCase() === 'retry-after'; });
      const value = key ? String(h[key]) : '';
      const retry = /^\d+$/.test(value) ? Number(value) * 1000 : Date.parse(value) - Date.now();
      pauseRetry_(Math.max(1000 * Math.pow(2, attempt), Number.isFinite(retry) ? retry : 0), deadline);
      continue;
    }
    throw new Error('API HTTP ' + status + (status === 401 ? '; reauthorize the affected account.' : '.'));
  }
  throw new Error('API retries exhausted.');
}

function pauseRetry_(ms, deadline) {
  if (Date.now() + ms + 1000 >= deadline) throw new Error('API retry exceeds sync time budget.');
  Utilities.sleep(ms);
}

function query_(params) {
  return Object.keys(params).filter(function (k) { return params[k] !== undefined; })
    .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); }).join('&');
}
