'use strict';

const assert = require('node:assert');
const { isInternalNav, isOpenableExternal } = require('./url-policy');

// isInternalNav
assert.strictEqual(isInternalNav('https://outlook.office.com/mail'), true);
assert.strictEqual(isInternalNav('https://login.microsoftonline.com/'), true);
assert.strictEqual(isInternalNav('https://example.com/x'), false);
assert.strictEqual(isInternalNav('file:///etc/passwd'), false);
assert.strictEqual(isInternalNav('https://us.azureauth.duosecurity.com/authorization'), true);

// isOpenableExternal
assert.strictEqual(isOpenableExternal('https://example.com/a?b=c'), true);
assert.strictEqual(isOpenableExternal('mailto:a@b.c'), true);
assert.strictEqual(isOpenableExternal('javascript:alert(1)'), false);
assert.strictEqual(isOpenableExternal('file:///etc/passwd'), false);

console.log('url-policy: all assertions passed');
