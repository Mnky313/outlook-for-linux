'use strict';

const assert = require('node:assert');
const { isInternalNav, isOpenableExternal, isAuthHost } = require('./url-policy');

// isInternalNav
assert.strictEqual(isInternalNav('https://outlook.office.com/mail'), true);
assert.strictEqual(isInternalNav('https://login.microsoftonline.com/'), true);
assert.strictEqual(isInternalNav('https://example.com/x'), false);
assert.strictEqual(isInternalNav('file:///etc/passwd'), false);
assert.strictEqual(isInternalNav('https://us.azureauth.duosecurity.com/authorization'), true);
assert.strictEqual(isInternalNav('https://api-abc123.duosecurity.com/frame'), true);
assert.strictEqual(isInternalNav('https://evil-duosecurity.com/x'), false);
assert.strictEqual(isInternalNav('https://login.microsoft.com/common/oauth2'), true);

// isAuthHost (auth-flow context detection)
assert.strictEqual(isAuthHost('https://login.microsoftonline.com/common/oauth2'), true);
assert.strictEqual(isAuthHost('https://device.login.microsoftonline.com/x'), true);
assert.strictEqual(isAuthHost('https://signup.live.com/signup'), true);
assert.strictEqual(isAuthHost('https://account.live.com/'), true);
assert.strictEqual(isAuthHost('https://api-abc123.duosecurity.com/frame'), true);
assert.strictEqual(isAuthHost('https://tenant.msauth.net/redirect'), true);
assert.strictEqual(isAuthHost('https://example.com/x'), false);
assert.strictEqual(isAuthHost('https://evil-microsoftonline.com/x'), false);
assert.strictEqual(isAuthHost('http://login.microsoftonline.com/'), false);
assert.strictEqual(isAuthHost('https://outlook.office.com/mail'), false);

// isOpenableExternal
assert.strictEqual(isOpenableExternal('https://example.com/a?b=c'), true);
assert.strictEqual(isOpenableExternal('mailto:a@b.c'), true);
assert.strictEqual(isOpenableExternal('javascript:alert(1)'), false);
assert.strictEqual(isOpenableExternal('file:///etc/passwd'), false);

console.log('url-policy: all assertions passed');
