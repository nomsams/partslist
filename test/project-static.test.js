const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const app = fs.readFileSync(path.join(root, 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

test('browser storage writes only public exchange-rate cache data', () => {
  const writes = [...app.matchAll(/localStorage\.setItem\(([^\n]+)/g)].map((match) => match[1]);
  assert.equal(writes.length, 1);
  assert.match(writes[0], /^cacheKey,/);
  assert.doesNotMatch(app, /localStorage\.setItem\([^\n]*(settings|warehouse)/i);
});

test('all DOM bindings resolve to unique element ids', () => {
  const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]);
  assert.equal(new Set(ids).size, ids.length);
  const references = [...app.matchAll(/el\('([^']+)'\)/g)].map((match) => match[1]);
  references.forEach((id) => assert.ok(ids.includes(id), `Missing DOM id: ${id}`));
});
