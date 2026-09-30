const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const I18n = require('../i18n.js');
const swedish = require('../i18n-sv.js');

I18n.register('sv', swedish);
const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8').replace(/\r\n/g, '\n');

const placeholders = (text) => [...text.matchAll(/\{(\d+)\}/g)].map((match) => match[1]).sort().join(',');

test('templates, plurals, sentences and lists are translated', () => {
  assert.equal(I18n.translate('sv', '1 part selected'), '1 artikel markerad');
  assert.equal(I18n.translate('sv', '3 parts selected'), '3 artiklar markerade');
  assert.equal(I18n.translate('sv', 'Landed in Sweden'), 'Ankommet till Sverige');
  // The status text the user reported as untranslated.
  assert.equal(I18n.translate('sv', 'Daily reference rates loaded for 2026-09-30. Source: Frankfurter.'), 'Dagliga referenskurser inlästa för 2026-09-30. Källa: Frankfurter.');
  assert.equal(
    I18n.translate('sv', 'Select the storage types you expect to pay for. Shelf racks with bins are enabled by default.'),
    'Välj de lagringstyper du räknar med att betala för. Hyllställ med lagerlådor är förvalt.',
  );
  assert.match(I18n.translate('sv', 'Sheet values use the Frakt column per line. Per line applies one amount to every line; Whole kit splits one amount across the kit’s stocked lines.'), /^Värden från bladet använder kolumnen Frakt per rad\./);
  // Several sentences, a list with a longer known phrase, and a title with a trailing full stop.
  assert.equal(I18n.translate('sv', 'Range cleared. Some currency rates are still missing.'), 'Området har rensats. Vissa valutakurser saknas fortfarande.');
  assert.equal(I18n.translate('sv', 'Portugal · EU · EEA agreement'), 'Portugal · EU · EES-avtalet');
  assert.equal(I18n.translate('sv', 'Apply for Tollkreditt.'), 'Ansök om Tollkreditt.');
  // Text that is not interface text passes through untouched, and English stays English.
  assert.equal(I18n.translate('sv', 'TE26778HT'), 'TE26778HT');
  assert.equal(I18n.translate('en', 'Import workbook'), 'Import workbook');
});

test('a template cannot swallow neighbouring segments of a " · " list', () => {
  // "{0} source prices" must translate only its own segment, not "Import (customs clearance) · USD source prices".
  assert.equal(
    I18n.translate('sv', 'United States → Norway · Import (customs clearance) · USD source prices'),
    'USA → Norge · Import (tullklarering) · inköpspriser i USD',
  );
  assert.equal(I18n.translate('sv', '5 shelf · 5 drawer · 1 pallet'), '5 hyllplats · 5 lådfack · 1 pall');
});

test('country names come from the language, and a template that changes nothing is not a translation', () => {
  assert.equal(I18n.countryName('sv', '840', 'United States'), 'USA');
  assert.equal(I18n.countryName('sv', '756', 'Switzerland'), 'Schweiz');
  assert.equal(I18n.countryName('sv', '100', 'Bulgaria'), 'Bulgarien');
  assert.equal(I18n.countryName('sv', '000', 'Somaliland'), 'Somaliland');
  assert.equal(I18n.translate('sv', 'United States → Norway'), 'USA → Norge');
  assert.equal(I18n.hasTranslation('sv', 'USD → SEK'), false);
});

test('every entry of the Swedish pack keeps its placeholders and no English key has two translations', () => {
  const seen = new Map();
  swedish.forEach(([english, translated]) => {
    assert.equal(placeholders(english), placeholders(translated), `Placeholders differ: ${english}`);
    assert.ok(translated.trim().length > 0, `Empty translation: ${english}`);
    if (seen.has(english) && seen.get(english) !== translated) assert.fail(`Conflicting translations for: ${english}`);
    seen.set(english, translated);
  });
  assert.ok(swedish.length > 900);
});

test('every static string in the page has a Swedish translation', () => {
  const body = html.replace(/<script[\s\S]*?<\/script>/g, '').replace(/<[^>]*translate="no"[^>]*>[\s\S]*?<\/[a-z]+>/g, '');
  const strings = new Set();
  const add = (raw) => {
    const text = raw.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();
    if (/[A-Za-z]{2}/.test(text)) strings.add(text);
  };
  for (const match of body.matchAll(/>([^<>]+)</g)) add(match[1]);
  for (const match of body.matchAll(/\b(?:title|placeholder|aria-label|data-tip|alt)="([^"]*)"/g)) add(match[1]);
  // Codes, units and names that read the same in both languages.
  const languageNeutral = new Set(['PL', 'AES-GCM', 'NOK', 'SEK', 'Language / Språk']);
  const missing = [...strings].filter((text) => !languageNeutral.has(text) && !I18n.hasTranslation('sv', text));
  assert.deepEqual(missing, []);
});
