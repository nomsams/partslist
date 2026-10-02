const test = require('node:test');
const assert = require('node:assert/strict');
const Vault = require('../vault.js');

const files = [
  { name: 'Parts list 2026.xlsx', bytes: new TextEncoder().encode('PK\u0003\u0004 pretend spreadsheet bytes '.repeat(50)) },
  { name: 'rates.json', bytes: new TextEncoder().encode('{"shelf":123}') },
];

test('documents survive sealing and opening with the same key', async () => {
  const sealed = await Vault.seal(files, 'a long key phrase', { iterations: 100000 });
  const opened = await Vault.open(sealed, 'a long key phrase');
  assert.deepEqual(opened.map((file) => file.name), ['Parts list 2026.xlsx', 'rates.json']);
  assert.deepEqual(opened[0].bytes, files[0].bytes);
  assert.deepEqual(opened[1].bytes, files[1].bytes);
});

test('a wrong key or a changed byte is rejected', async () => {
  const sealed = await Vault.seal(files, 'a long key phrase', { iterations: 100000 });
  await assert.rejects(() => Vault.open(sealed, 'another key phrase'), /key is wrong/);
  const tampered = sealed.slice();
  tampered[tampered.length - 5] ^= 1;
  await assert.rejects(() => Vault.open(tampered, 'a long key phrase'), /key is wrong/);
  await assert.rejects(() => Vault.open(new Uint8Array(10), 'a long key phrase'), /damaged/);
});

test('the sealed file reveals nothing readable: no names, no spreadsheet signature, different every time', async () => {
  const first = await Vault.seal(files, 'a long key phrase', { iterations: 100000 });
  const second = await Vault.seal(files, 'a long key phrase', { iterations: 100000 });
  const text = Buffer.from(first).toString('latin1');
  ['Parts list', 'xlsx', 'rates', 'PK', 'shelf'].forEach((word) => assert.ok(!text.includes(word), `leaks ${word}`));
  assert.notDeepEqual(first, second);                 // fresh salt and IV each time
});
