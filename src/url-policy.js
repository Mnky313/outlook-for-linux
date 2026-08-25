'use strict';

const ALLOWED_NAV_HOSTS = [
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.live.com',
  'login.microsoftonline.com',
  'login.live.com',
  'signup.live.com',
];

const OPENABLE_PROTOCOLS = new Set(['https:', 'http:', 'mailto:']);

function isInternalNav(urlStr) {
  try {
    const u = new URL(urlStr);
    return u.protocol === 'https:' && ALLOWED_NAV_HOSTS.includes(u.hostname);
  } catch { return false; }
}

function isOpenableExternal(urlStr) {
  try {
    const u = new URL(urlStr);
    return OPENABLE_PROTOCOLS.has(u.protocol);
  } catch { return false; }
}

module.exports = { isInternalNav, isOpenableExternal };
