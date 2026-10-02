// Packs documents into one encrypted file that is safe to commit: data/bundle.dat
//
//   PARTSLIST_KEY="your key phrase" node scripts/encrypt-documents.js "parts list.xlsx" "price list.xlsx" rates.json
//
// The key is read from the environment so it is never written into the repository or the code.
// The app opens the bundle when the same key is typed into "Unlock saved documents".
const fs = require('node:fs');
const path = require('node:path');
const Vault = require('../vault.js');

async function main() {
  const key = process.env.PARTSLIST_KEY;
  const inputs = process.argv.slice(2);
  if (!key || !inputs.length) {
    process.stderr.write('Usage: PARTSLIST_KEY="key phrase" node scripts/encrypt-documents.js <file> [<file> ...]\n');
    process.exit(1);
  }
  const files = inputs.map((input) => ({ name: path.basename(input), bytes: new Uint8Array(fs.readFileSync(input)) }));
  const sealed = await Vault.seal(files, key);
  const target = path.resolve(__dirname, '..', 'data', 'bundle.dat');
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, sealed);
  // Prove it opens before anything is committed.
  const check = await Vault.open(new Uint8Array(fs.readFileSync(target)), key);
  process.stdout.write(`Wrote ${path.relative(process.cwd(), target)} (${sealed.length} bytes) with ${check.length} document(s).\n`);
}

main().catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exit(1);
});
