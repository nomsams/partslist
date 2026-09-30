const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const app = fs.readFileSync(path.join(root, 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const css = fs.readFileSync(path.join(root, 'styles.css'), 'utf8');

test('browser storage writes only public data: exchange-rate cache and interface preferences', () => {
  const writes = [...app.matchAll(/localStorage\.setItem\(([^\n]+)/g)].map((match) => match[1]);
  assert.deepEqual(writes.map((write) => write.split(',')[0]).sort(), ['UI_PREFERENCES_KEY', 'cacheKey']);
  // The preferences record holds nothing but the language and whether the settings panel is collapsed.
  assert.match(app, /JSON\.stringify\(\{ language: state\.language, sidebarCollapsed: state\.sidebarCollapsed \}\)/);
  assert.doesNotMatch(app, /localStorage\.setItem\([^\n]*(settings|warehouse)/i);
});

test('all DOM bindings resolve to unique element ids', () => {
  const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]);
  assert.equal(new Set(ids).size, ids.length);
  const references = [...app.matchAll(/el\('([^']+)'\)/g)].map((match) => match[1]);
  references.forEach((id) => assert.ok(ids.includes(id), `Missing DOM id: ${id}`));
});

test('non-technical users have guided setup, terminology help, and save-state feedback', () => {
  assert.match(html, /id="setup-guide"/);
  assert.match(html, /id="setup-checklist"/);
  assert.match(html, /What do these terms mean\?/);
  assert.match(html, /Landed cost/);
  assert.match(html, /id="save-state"/);
  assert.match(app, /function renderSetupGuide\(/);
  assert.match(app, /function renderSaveState\(/);
});

test('the consolidated list has an isolated opaque sticky table and 3PL rate import is discoverable', () => {
  assert.match(html, /aria-label="Scrollable consolidated parts table"/);
  assert.match(css, /#consolidated-view\s*\{[^}]*overflow:\s*auto/s);
  assert.match(css, /#consolidated-table\s*\{[^}]*overflow:\s*auto/s);
  assert.match(css, /#consolidated-table\s*\{[^}]*isolation:\s*isolate/s);
  assert.match(css, /\.parts-table th\.select-cell\s*\{[^}]*z-index:\s*30/s);
  assert.match(css, /\.parts-table \.select-cell\s*\{[^}]*background-color:\s*#fff/s);
  assert.match(css, /grid-template-rows:\s*auto auto minmax\(0, 1fr\)/);
  assert.match(html, />Import 3PL rates JSON</);
  assert.match(html, /id="warehouse-rate-editor"/);
});

test('Swedish is the default and additional workbooks can append routed kits', () => {
  assert.match(html, /<html lang="sv">/);
  assert.match(html, /id="language-select"[^>]*role="group"/);
  assert.match(html, /data-language="sv" aria-pressed="true"/);
  assert.match(html, /data-language="en" aria-pressed="false"/);
  assert.match(html, /src="i18n\.js"[\s\S]*src="i18n-sv\.js"[\s\S]*src="app\.js"/);
  assert.match(html, /id="add-kit-file"[^>]*multiple/);
  assert.match(html, /id="warehouse-country"/);
  assert.match(app, /function addKitWorkbooks\(/);
  assert.match(app, /manufacturer:\s*''/);
  assert.match(app, /originCountry:\s*'840'/);
  // Kit discount, source currency and duty default to null, meaning "use the project-wide value".
  assert.match(app, /defaultCurrency:\s*null/);
  assert.match(app, /discountRate:\s*null/);
  assert.match(app, /dutyRate:\s*null/);
  assert.match(app, /function kitDiscount\(/);
  assert.match(app, /Core\.customsTreatment\(/);
  assert.match(app, /manufacturer:\s*state\.supplier\.name/);
  assert.match(app, /function enabledKitRoutes\(/);
});

test('settings sidebar can be collapsed and warehouse rate files are password-free', () => {
  assert.match(html, /id="sidebar-toggle"[^>]*aria-controls="app-sidebar"/);
  assert.match(html, /id="app-sidebar"/);
  assert.match(app, /function renderSidebarState\(/);
  // Collapsing shrinks the panel to a slim rail instead of removing it, and hides the body from keyboards.
  assert.match(html, /id="sidebar-collapse"/);
  assert.match(html, /id="sidebar-expand"/);
  assert.match(css, /\.workspace\.sidebar-collapsed\s*\{[^}]*grid-template-columns:\s*44px/);
  assert.match(css, /\.workspace\.sidebar-collapsed \.sidebar-body\s*\{\s*display:\s*none/);
  assert.match(app, /toggleAttribute\('inert', collapsed\)/);
  assert.match(html, /No password is required for the rates JSON/);
});

test('isometric rack paints front uprights after shelves so the nearest support stays visible', () => {
  const warehouseRenderer = app.slice(app.indexOf('function buildWarehouseIsoSvg'), app.indexOf('function formatNok'));
  const shelfLoop = warehouseRenderer.indexOf('for (let level = 0; level < rackLevels; level += 1)');
  const frontUprights = warehouseRenderer.indexOf("'rack-upright rack-upright-front'");
  assert.ok(shelfLoop >= 0 && frontUprights > shelfLoop, 'Front rack uprights must be painted after the shelves');
  assert.match(warehouseRenderer, /post <= rackBays/);
});

test('warehouse storage types are explicit and shelves are the default', () => {
  assert.match(html, /id="wh-shelf-enabled"[^>]*checked/);
  assert.match(html, /id="wh-drawer-enabled"/);
  assert.match(html, /id="wh-pallet-enabled"/);
  assert.match(html, /id="wh-planned-pallets"/);
  assert.match(app, /plannedPallets:\s*0/);
  assert.match(app, /Core\.planWarehouseStorage\(/);
});

test('normal Excel export is a re-importable whole-project workbook', () => {
  assert.match(html, />Save project \.xlsx</);
  assert.match(app, /function buildProjectDataSheet\(/);
  assert.match(app, /mappings,\s*\n\s*variables:/);
  assert.match(app, /wireRelations: state\.wireRelations/);
  assert.match(app, /rates: \{ \.\.\.state\.warehouse\.rates \}/);
  assert.match(app, /-partslist-project\.xlsx/);
});

test('routine status text does not consume a permanent row', () => {
  assert.match(css, /\.status-bar:not\(\.busy\):not\(\.error\):not\(\.warning\)\s*\{\s*display:\s*none/);
  assert.match(css, /\.app-header\s*\{[^}]*min-height:\s*54px/s);
});

test('every header control shares one height and style, and the customer country is its own setting', () => {
  assert.match(css, /\.app-header \.header-button,\s*\.app-header \.segmented\s*\{\s*height:\s*34px/);
  const headerButtons = html.slice(html.indexOf('<header'), html.indexOf('</header>')).match(/class="button header-button[^"]*"/g) || [];
  assert.ok(headerButtons.length >= 5);
  assert.doesNotMatch(html.slice(html.indexOf('<header'), html.indexOf('</header>')), /button (ghost|primary)/);
  assert.match(html, /data-bind="customer\.country" data-customer-select/);
  assert.match(app, /customer: \{ country: '578' \}/);
  assert.match(app, /function isCrossBorderSale\(/);
});
