(function exposePartsListVault(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.PartsListVault = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  'use strict';

  // An encrypted bundle of documents (parts lists, rate files) that can live in a public repository.
  // The file has no readable header: salt, iteration count, IV, then AES-256-GCM ciphertext of a gzipped
  // container. File names and types are inside the ciphertext, so the committed file says nothing about its content.
  const SALT_BYTES = 16;
  const IV_BYTES = 12;
  const DEFAULT_ITERATIONS = 600000;
  const MIN_ITERATIONS = 100000;
  const MAX_ITERATIONS = 2000000;
  const HEADER_BYTES = SALT_BYTES + 4 + IV_BYTES;

  const subtle = () => globalThis.crypto.subtle;

  async function deriveKey(passphrase, salt, iterations, usages) {
    const material = await subtle().importKey('raw', new TextEncoder().encode(String(passphrase).normalize('NFKC')), 'PBKDF2', false, ['deriveKey']);
    return subtle().deriveKey({ name: 'PBKDF2', salt, iterations, hash: 'SHA-256' }, material, { name: 'AES-GCM', length: 256 }, false, usages);
  }

  async function pipe(bytes, transform) {
    const stream = new Blob([bytes]).stream().pipeThrough(transform);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  // Container: 4-byte length of a JSON list of { name, size }, that list, then every file's bytes in order.
  function pack(files) {
    const list = files.map((file) => ({ name: String(file.name), size: file.bytes.length }));
    const header = new TextEncoder().encode(JSON.stringify(list));
    const total = 4 + header.length + files.reduce((sum, file) => sum + file.bytes.length, 0);
    const out = new Uint8Array(total);
    new DataView(out.buffer).setUint32(0, header.length);
    out.set(header, 4);
    let offset = 4 + header.length;
    files.forEach((file) => { out.set(file.bytes, offset); offset += file.bytes.length; });
    return out;
  }

  function unpack(bytes) {
    const headerLength = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0);
    if (headerLength <= 0 || 4 + headerLength > bytes.length) throw new Error('The document bundle is damaged.');
    const list = JSON.parse(new TextDecoder().decode(bytes.subarray(4, 4 + headerLength)));
    if (!Array.isArray(list)) throw new Error('The document bundle is damaged.');
    let offset = 4 + headerLength;
    return list.map((entry) => {
      const size = Number(entry.size);
      if (!Number.isInteger(size) || size < 0 || offset + size > bytes.length) throw new Error('The document bundle is damaged.');
      const file = { name: String(entry.name), bytes: bytes.slice(offset, offset + size) };
      offset += size;
      return file;
    });
  }

  async function seal(files, passphrase, { iterations = DEFAULT_ITERATIONS } = {}) {
    if (!String(passphrase || '').length) throw new Error('A key is required.');
    const salt = globalThis.crypto.getRandomValues(new Uint8Array(SALT_BYTES));
    const iv = globalThis.crypto.getRandomValues(new Uint8Array(IV_BYTES));
    const key = await deriveKey(passphrase, salt, iterations, ['encrypt']);
    const packed = await pipe(pack(files), new CompressionStream('gzip'));
    const ciphertext = new Uint8Array(await subtle().encrypt({ name: 'AES-GCM', iv }, key, packed));
    const out = new Uint8Array(HEADER_BYTES + ciphertext.length);
    out.set(salt, 0);
    new DataView(out.buffer).setUint32(SALT_BYTES, iterations);
    out.set(iv, SALT_BYTES + 4);
    out.set(ciphertext, HEADER_BYTES);
    return out;
  }

  async function open(bytes, passphrase) {
    if (!(bytes instanceof Uint8Array) || bytes.length < HEADER_BYTES + 17) throw new Error('The document bundle is damaged.');
    const iterations = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(SALT_BYTES);
    if (iterations < MIN_ITERATIONS || iterations > MAX_ITERATIONS) throw new Error('The document bundle is damaged.');
    const salt = bytes.slice(0, SALT_BYTES);
    const iv = bytes.slice(SALT_BYTES + 4, HEADER_BYTES);
    const key = await deriveKey(passphrase, salt, iterations, ['decrypt']);
    let packed;
    try {
      packed = new Uint8Array(await subtle().decrypt({ name: 'AES-GCM', iv }, key, bytes.slice(HEADER_BYTES)));
    } catch {
      throw new Error('The key is wrong, or the document bundle is damaged.');
    }
    return unpack(await pipe(packed, new DecompressionStream('gzip')));
  }

  return { seal, open, DEFAULT_ITERATIONS };
}));
