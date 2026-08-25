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
// api-*.duosecurity.com); allow the vendor's auth domain by suffix instead of
// chasing exact hostnames. Pattern borrowed from teams-for-linux.
const ALLOWED_NAV_HOST_SUFFIXES = ['.duosecurity.com'];

function isAllowedHost(hostname) {
  return (
    ALLOWED_NAV_HOSTS.includes(hostname) ||
    ALLOWED_NAV_HOST_SUFFIXES.some((s) => hostname.endsWith(s))
  );
}

const OPENABLE_PROTOCOLS = new Set(['https:', 'http:', 'mailto:']);

function isInternalNav(urlStr) {
  try {
    const u = new URL(urlStr);
    return u.protocol === 'https:' && isAllowedHost(u.hostname);
  } catch { return false; }
}

function isOpenableExternal(urlStr) {
  try {
    const u = new URL(urlStr);
    return OPENABLE_PROTOCOLS.has(u.protocol);
  } catch { return false; }
}

module.exports = { isInternalNav, isOpenableExternal };
