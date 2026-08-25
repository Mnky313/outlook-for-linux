'use strict';

const ALLOWED_NAV_HOSTS = [
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.live.com',
  'login.microsoftonline.com',
  'login.microsoft.com',
  'login.live.com',
  'signup.live.com',
];

// Institutional 2FA (Duo) bounces across per-tenant hosts (us.azureauth.*,
// api-*.duosecurity.com); allow the vendor's auth domain by suffix.
const ALLOWED_NAV_HOST_SUFFIXES = ['.duosecurity.com'];
// Microsoft is migrating the Office app shells (OWA, m365 portal) to
// *.cloud.microsoft — same category as outlook.office.com.
const ALLOWED_NAV_APP_SUFFIXES = ['.cloud.microsoft'];

const OPENABLE_PROTOCOLS = new Set(['https:', 'http:', 'mailto:']);

// Identity-provider infrastructure. No static list can enumerate every host an
// auth flow touches (signup, device login, per-tenant endpoints), so instead of
// allowlisting targets we detect being INSIDE an auth flow: main.js allows
// any https navigation/popups while the current page is an auth host, and
// routes links out everywhere else (reading mail).
const AUTH_HOSTS = [
  'login.microsoftonline.com',
  'login.microsoft.com',
  'login.live.com',
  'signup.live.com',
  'account.live.com',
];
const AUTH_HOST_SUFFIXES = [
  '.microsoftonline.com',
  '.duosecurity.com',
  '.msauth.net',
  '.msftauth.net',
  '.msidentity.com',
];

function isAuthHost(urlStr) {
  try {
    const u = new URL(urlStr);
    return (
      u.protocol === 'https:' &&
      (AUTH_HOSTS.includes(u.hostname) ||
        AUTH_HOST_SUFFIXES.some((s) => u.hostname.endsWith(s)))
    );
  } catch {
    return false;
  }
}

function isInternalNav(urlStr) {
  try {
    const u = new URL(urlStr);
    return (
      u.protocol === 'https:' &&
      (ALLOWED_NAV_HOSTS.includes(u.hostname) ||
        [...ALLOWED_NAV_HOST_SUFFIXES, ...ALLOWED_NAV_APP_SUFFIXES].some((s) =>
          u.hostname.endsWith(s)
        ))
    );
  } catch {
    return false;
  }
}

function isOpenableExternal(urlStr) {
  try {
    const u = new URL(urlStr);
    return OPENABLE_PROTOCOLS.has(u.protocol);
  } catch { return false; }
}

module.exports = { isInternalNav, isOpenableExternal, isAuthHost };
