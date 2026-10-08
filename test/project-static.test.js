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
  // The preferences record holds only harmless interface choices: language, whether the settings panel is
  // collapsed, and how the parts table is arranged (sort column and hidden columns). No business data.
  assert.match(app, /language: state\.language,/);
  assert.match(app, /sidebarCollapsed: state\.sidebarCollapsed,/);
  assert.match(app, /consolidatedSort: state\.consolidatedSort,/);
  assert.match(app, /hiddenColumns: state\.hiddenColumns,/);
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
  // A kit field left null means "use the project-wide value"; check the KIT_DEFAULTS record directly.
  const kitDefaults = app.match(/KIT_DEFAULTS = Object\.freeze\(\{([^}]*)\}/)[1];
  assert.match(kitDefaults, /defaultCurrency:\s*null/);
  assert.match(kitDefaults, /discountRate:\s*null/);
  assert.match(kitDefaults, /dutyRate:\s*null/);
  assert.match(app, /function kitDiscount\(/);
  assert.match(app, /Core\.customsTreatment\(/);
  // An added kit falls back to the supplier as manufacturer unless the file name overrides it.
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
  // Front uprights are identified by their CSS class, not by the exact text of the loop that draws them.
  const frontUprights = warehouseRenderer.indexOf('rack-upright-front');
  // The bin loop that paints the shelf decks and bins must come first.
  const shelfBins = warehouseRenderer.indexOf('locationIndex * binsPerLocation');
  // The rack is taller than its top bins, so no beam or post is painted across a bin: every level leaves room
  // for the bins under the next deck and the top beams sit above the top level.
  assert.match(warehouseRenderer, /rackHeight = 5 \+ rackLevels \* levelHeight/);
  assert.match(warehouseRenderer, /binHeight = \(levelHeight - deckHeight - 4 - binGap \* \(high - 1\)\) \/ high/);
  // The numbers in the label are the bins in use and the bins available, not the capacity twice.
  assert.match(warehouseRenderer, /BINS \$\{Math\.round\(model\.shelfBins\)\} USED OF \$\{model\.totalBins\}/);
  assert.ok(shelfBins >= 0, 'The renderer must draw shelf bins');
  assert.ok(frontUprights > shelfBins, 'Front rack uprights must be painted after the shelves');
  // One upright per post across the rack width.
  // Equivalent loop spellings are fine; what matters is one upright per post across the rack width.
  assert.match(warehouseRenderer, /post\s*<=\s*rackBays/);
  assert.match(warehouseRenderer, /rack-upright rack-upright-front/);
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

test('VAT is off and locked by default, and only applies where a sales country is set to "price"', () => {
  // Read the values out of the source rather than matching their exact source formatting.
  const vatRate = Number(app.match(/vatRate:\s*([0-9.]+)/)[1]);
  assert.equal(vatRate, 0, 'VAT must start at zero');
  const vatDefaults = app.match(/VAT_DEFAULTS = Object\.freeze\(\{([^}]*)\}/)[1];
  assert.match(vatDefaults, /578:\s*'none'/);
  assert.match(vatDefaults, /752:\s*'price'/);
  assert.match(app, /unlocked:\s*false/);
  assert.match(app, /Core\.salesVatRate\(/);
  // Every VAT rate field starts read-only and is opened by a padlock.
  const rateFields = html.match(/<input[^>]*data-bind="config\.vatRate"[^>]*>/g) || [];
  assert.equal(rateFields.length, 3);
  rateFields.forEach((field) => {
    assert.match(field, /data-vat-field/);
    assert.match(field, /readonly/);
    assert.match(field, /value="0"|placeholder=/);
  });
  assert.equal((html.match(/data-vat-lock/g) || []).length, 3);
  // Generated sheets from a previous export are not listed as source sheets.
  assert.match(app, /filter\(\(sheetName\) => !isGeneratedSheet\(state\.workbook\.Sheets\[sheetName\]\)\)\.map\(\(sheetName\) => createTab/);
});

test('the map carries the editable tables on its route: one arrow through shipment & tolls, one north through customer pricing', () => {
  assert.match(html, /id="map-stage"/);
  assert.match(html, /id="map-overlay"/);
  ['warehouse', 'customer'].forEach((name) => {
    assert.match(html, new RegExp(`id="flow-box-${name}"`));
    assert.match(html, new RegExp(`id="flow-leg-${name}"`));
  });
  // Every manufacturer gets its own lane and its own shipment table; only the warehouse side is shared.
  assert.match(html, /id="flow-lanes"/);
  assert.match(app, /function buildMakerLane\(/);
  assert.match(app, /function buildMakerBox\(/);
  // The tables move between the route row (Route view) and the map (Map view); wires are drawn up to a table
  // and on from its far side, and the tables are placed clear of the map markers.
  // The flow table moves to the map overlay on wide screens and back onto the route row when narrow;
  // both destinations must be referenced somewhere in the renderer.
  assert.match(app, /dom\.mapBoxes[\s\S]{0,200}append\(refs\.box\)/);
  assert.match(app, /function renderMapScene\(/);
  assert.match(app, /Core\.placeBox\(/);
  assert.match(app, /Core\.curveAroundRect\(/);
  assert.match(css, /\.map-boxes \.flow-box\s*\{[^}]*position:\s*absolute/s);
  // Narrow windows stack the two panels and list the tables under the map instead of on it.
  assert.match(css, /\.map-stage\.stacked \.map-panels/);
  assert.match(css, /\.map-boxes\.flow\s*\{[^}]*position:\s*static/s);
});

test('a multi-kit catalogue file is split into one kit sheet per kit, from Bulgaria, with net prices and no extra discount', () => {
  assert.match(app, /function catalogueSheetsOf\(/);
  assert.match(app, /Core\.parseKitCatalogue\(/);
  assert.match(app, /Core\.catalogueSheetRows\(/);
  assert.match(app, /CATALOGUE_DEFAULTS = Object\.freeze\(\{ manufacturer: 'Häny', originCountry: '100', freightPercent: 3 \}\)/);
  assert.match(app, /discountRate: 0,/);
  // Opening a catalogue as the first file, or adding one to an open project, takes the same route.
  assert.match(app, /function openWorkbookFile\(/);
  assert.match(app, /if \(file\) openWorkbookFile\(file\)/);
  // Several manufacturers in one warehouse: a second article number column and per-manufacturer filters.
  assert.match(app, /altPart: 'Alt\. part no\.'/);
  assert.match(app, /makercommon:\$\{maker\}/);
});

test('saved documents are an encrypted bundle that the start screen can unlock, and readable copies are ignored by git', () => {
  assert.match(html, /id="vault-form"[^>]*hidden|class="vault-form hidden"/);
  assert.match(html, /src="vault\.js"[\s\S]*src="app\.js"/);
  assert.match(app, /VAULT_FILE = 'data\/bundle\.dat'/);
  assert.match(app, /PartsListVault\.open\(/);
  const ignore = fs.readFileSync(path.join(root, '.gitignore'), 'utf8');
  ['*.xlsx', '*.partslist', 'warehouse-rates.private.json'].forEach((pattern) => assert.ok(ignore.includes(pattern), `${pattern} must be ignored`));
  // The committed bundle must not leak what is inside it, and the key must not be written anywhere in the repository.
  const bundle = fs.readFileSync(path.join(root, 'data', 'bundle.dat')).toString('latin1');
  ['xlsx', 'Valve', 'Spare', 'Minimal', 'PART NUMBER'].forEach((word) => assert.ok(!bundle.includes(word), `bundle leaks ${word}`));
});

test('ease of use: column presets, price explainer, problem bar, customer price list, setup dialog and example data', () => {
  ['simple', 'purchase', 'sales', 'full'].forEach((name) => assert.match(html, new RegExp(`data-column-preset="${name}"`)));
  assert.match(html, /id="problem-bar"/);
  assert.match(html, /id="explain-panel"/);
  assert.match(html, /id="maker-setup"/);
  assert.match(html, /id="customer-export"/);
  assert.match(html, /id="example-data"/);
  assert.match(app, /function explainSteps\(/);
  assert.match(app, /function collectProblems\(/);
  assert.match(app, /function exportCustomerPriceList\(/);
  assert.match(app, /function openMakerSetup\(/);
  assert.match(app, /function resetMaker\(/);
  // The customer list must never carry costs.
  const customer = app.slice(app.indexOf('function exportCustomerPriceList('), app.indexOf('/* -- A short setup dialog'));
  assert.doesNotMatch(customer, /landedCost|discountedTotal|kitFreight|fxMargin/);
  // The saved documents open without the setup dialog; a file added by hand gets it.
  assert.match(app, /addKitWorkbooks\(lists, \{ review: false, quiet: true \}\)/);
});

test('popularity: sales histories per market, popular-in-both, kit ranking and suggested stock', () => {
  assert.match(html, /id="popularity-view"/);
  ['parts', 'both', 'kits', 'stock'].forEach((mode) => assert.match(html, new RegExp(`data-pop-mode="${mode}"`)));
  assert.match(html, /id="popularity-file"[^>]*multiple/);
  assert.match(html, /data-column-preset="popularity"/);
  assert.match(app, /VIEW = Object\.freeze\(\{[^}]*popularity/);
  ['parseSalesHistory', 'buildPopularity', 'suggestStock'].forEach((name) => assert.ok(app.includes(`Core.${name}(`), name));
  // The sales history is business data: part of the saved project, never of browser storage.
  const prefs = app.slice(app.indexOf('function saveUiPreferences('), app.indexOf('function purgeLegacyLocalBusinessData('));
  assert.doesNotMatch(prefs, /history/);
  assert.match(app, /history: \{\s*years: state\.history\.years/);
  // Histories chosen through the normal import buttons are recognised, and the saved documents may hold them.
  assert.ok(app.includes('splitHistoryFiles('));
  // Storage: one article per bin and bins stacked high and deep are warehouse settings that the stock plan uses.
  assert.ok(html.includes('id="wh-one-per-bin"') && html.includes('id="wh-bins-high"') && html.includes('id="wh-bins-deep"'));
  assert.ok(app.includes('function planStorageFor(') && app.includes('function storageNeeds('));
});

test('human-machine interface: readable text, reachable targets, keyboard and screen-reader support, plain-language help', () => {
  const css = fs.readFileSync(path.join(root, 'styles.css'), 'utf8');
  // No ordinary text below 12 px: only the compact map boxes and drawings may be smaller.
  const small = [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)]
    .filter(([, selector, body]) => /font-size:\s*(?:[0-9]|1[01])(?:\.\d+)?px/.test(body) && !/\.map-|\.mini-|\.flow-|\.wire-|\.viz-|\.pop-bar|\bsvg\b|\.warehouse-preview|\.rack|\.kit-quick|\.kit-chip/.test(selector));
  assert.deepEqual(small.map(([, selector]) => selector.trim()), []);
  assert.ok(css.includes(':where(a, button, input, select, textarea, summary, [tabindex]):focus-visible'));
  assert.match(css, /prefers-reduced-motion: reduce/);
  assert.match(css, /\.skip-link/);
  assert.match(html, /class="skip-link"/);
  assert.match(html, /id="view-tabs"[^>]*role="tablist"/);
  assert.ok(app.includes("setAttribute('role', 'tab')"));
  assert.ok(app.includes('function handleTabKeydown('));
  // Messages stay long enough to read, errors longer, and a click dismisses them.
  assert.ok(app.includes('function toastDuration(') && app.includes('function bindToast('));
  // Every column has a one-sentence explanation.
  assert.ok(app.includes('const COLUMN_HELP = Object.freeze('));
  assert.ok(app.includes('const TAB_HELP = Object.freeze('));
  assert.ok(app.includes('needs-save'));
});

test('minimal parts lists: paired with the full list, compared side by side, switchable, common parts list with percentages', () => {
  assert.match(html, /id="minimal-view"/);
  assert.match(html, /data-parts-mode="full"/);
  assert.match(html, /data-parts-mode="minimal"/);
  assert.match(html, /id="minimal-file"[^>]*multiple/);
  assert.match(html, /id="common-mode"/);
  assert.match(html, /id="common-summary"/);
  ['minimalSheetFor', 'isActiveSheet', 'isHiddenByMode', 'addMinimalLists', 'compareRig', 'setPartsMode', 'renderCommonSummary', 'inheritedFreight'].forEach((name) => assert.ok(app.includes(`function ${name}(`), name));
  // The parts mode and the pairing travel with the saved project.
  assert.ok(app.includes('partsMode: state.partsMode'));
  assert.ok(app.includes('minimalOf'));
  // Everything that reads the sheets in play goes through isActiveSheet, not the raw enabled flag.
  assert.ok(!app.includes('state.mappings[name]?.enabled'));
  // The saved documents add the minimal lists after the full lists.
  assert.ok(app.includes('addMinimalLists(minimalLists'));
  // The README gives no supplier names or item numbers (example names only).
  const readme = fs.readFileSync(path.join(root, 'README.md'), 'utf8');
  assert.doesNotMatch(readme, /TEI|Häny|ET-PAK|Rock Drills/);
});

test('currency add-ons are global per currency, discount and duty per manufacturer, warehouse fields follow the storage types', () => {
  ['USD', 'EUR', 'NOK', 'DKK'].forEach((code) => assert.match(html, new RegExp(`data-fx-row="${code}"`)));
  assert.ok(app.includes("const FX_ADDON_CURRENCIES = Object.freeze(['USD', 'EUR', 'NOK', 'DKK'])"));
  // The add-on comes from the currency, not from the kit.
  assert.ok(/function kitFxMargin\(kit\) \{\s*return fxAddonFor\(kitCurrency\(kit\)\);/.test(app));
  assert.ok(app.includes('fxAddons: { ...state.fxAddons }'));
  // The kit card has no discount, duty or buffer fields of its own.
  const kitCard = app.slice(app.indexOf('function createKitFreightControl('), app.indexOf('// Discount and customs duty belong to the manufacturer'));
  assert.doesNotMatch(kitCard, /discountRate|dutyRate|'fxMargin'/);
  assert.ok(app.includes('function unifyMakerRates('));
  // One share on shelves; the drawers get the rest.
  assert.ok(app.includes('function normalizeStorageShares('));
  assert.ok(html.includes('data-show-when="shelf+drawer"') && html.includes('data-show-when="pallet"') && html.includes('data-show-when="drawer"'));
  assert.ok(app.includes('function applyWarehouseVisibility('));
  // Presets for the monthly activity come from the sales history.
  assert.match(html, /id="activity-preset"/);
  assert.ok(app.includes('function computeActivity(') && app.includes('ACTIVITY_PRESETS'));
});

test('canvas: grouped compact tables, flowing wires in and out, editing, and one encrypted workspace file', () => {
  assert.match(html, /id="canvas-view"/);
  ['cv-stage', 'cv-world', 'cv-wires', 'cv-inspector', 'cv-edit', 'cv-fit', 'cv-save', 'cv-open', 'cv-fullscreen', 'workspace-dialog'].forEach((id) => assert.ok(html.includes(`id="${id}"`), id));
  assert.ok(html.includes('id="cv-arrow-in"') && html.includes('id="cv-arrow-out"'));
  ['buildCanvasModel', 'cvNeighbours', 'cvResolve', 'cvDraw', 'cvCommit', 'cvStartEdit', 'openWorkspaceDialog', 'runWorkspace', 'bindCanvas'].forEach((name) => assert.ok(app.includes(`function ${name}(`), name));
  assert.match(app, /VIEW = Object\.freeze\(\{[^}]*canvas/);
  // The list tables show the starting prices only (no price after discount); the start table has the prices after each step when opened.
  assert.ok(app.includes("const CV_SHEET_COMPACT = Object.freeze(['part', 'name', 'qty', 'price'])"));
  assert.ok(!/CV_SHEET_(?:COMPACT|OPEN) = Object\.freeze\([^)]*'after'/.test(app));
  assert.ok(app.includes("const CV_CHAIN = Object.freeze(['start', 'disc', 'frt', 'mult', 'wh', 'out', 'cust'])"));
  ['cvStepCards', 'cvLayout', 'cvPipes', 'cvCustomerUnit'].forEach((name) => assert.ok(app.includes(`function ${name}(`), name));
  // The incoming freight says plainly what it is based on, and any card can be dragged from anywhere.
  assert.ok(app.includes("'of the discounted price, for each item'"));
  assert.ok(app.includes("event.target.closest('.cv-card')"));
  // The customer step is optional: nothing is added unless it is switched on.
  assert.ok(app.includes('function defaultOutgoing()') && app.includes('outgoing: state.outgoing'));
  assert.match(html, /data-parts-mode="minimal"[^>]*>Minimal kits/);
  // The layout and the view are part of the saved settings, so a workspace file restores the canvas as it was left.
  assert.ok(app.includes('canvas: { positions: state.canvas.positions'));
  assert.ok(app.includes('state.canvas = restoreCanvas(settings.canvas)'));
  // The workspace file is the encrypted secure project, never a readable file.
  const workspace = app.slice(app.indexOf('async function runWorkspace('), app.indexOf('/* ---------- Sales history & popularity'));
  assert.ok(workspace.includes('exportSecureProject()') && workspace.includes('importSecureProject(file)'));
  // Flowing wires move, but not for people who ask for less motion.
  const css = fs.readFileSync(path.join(root, 'styles.css'), 'utf8');
  assert.match(css, /@keyframes cv-flow/);
  assert.match(css, /prefers-reduced-motion: reduce/);
});

test('every manufacturer is its own route: own settings, exchange-rate buffer, percentage freight and an independent shipment', () => {
  assert.match(html, /id="makers-panel"/);
  assert.match(app, /function buildMakerCard\(/);
  assert.match(app, /MAKER_KIT_FIELDS = new Set\(/);
  assert.match(app, /function effectiveFx\(/);
  assert.match(app, /percent: 'Percent of item value'/);
  // Freight pools: a manufacturer that ships on its own does not share the project shipment.
  assert.match(app, /makerHasOwnShipment\(item\.manufacturer\)/);
  assert.doesNotMatch(app, /key: 'amountUsd'/);
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
