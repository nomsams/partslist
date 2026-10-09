const test = require('node:test');
const assert = require('node:assert/strict');
const {
  assessProfitabilityCompleteness,
  buildPopularity,
  parseSalesHistory,
  suggestStock,
  binsForUnits,
  analyzeBins,
  planWarehouseStorage: planStorage,
  allocationShares,
  EU_MEMBERS,
  catalogueSheetName,
  catalogueSheetRows,
  clampWirePosition,
  parseKitCatalogue,
  curveAroundRect,
  hasMissingInputs,
  missingInputReasons,
  sortItems,
  fitProjection,
  placeBox,
  quadraticBow,
  rayToRect,
  rectAround,
  rectContains,
  rectsOverlap,
  customsTreatment,
  importBounds,
  normalizeCountryCode,
  planWarehouseStorage,
  projectScenario,
  requiredLocations,
  salesVatRate,
  sanitizeWirePositions,
  summarizeOccurrences,
} = require('../partslist-core.js');

// The helpers are destructured above, so the tests below call them directly.

test('country codes retain valid ISO numeric values and reject malformed values', () => {
  assert.equal(normalizeCountryCode('752'), '752');
  assert.equal(normalizeCountryCode(276), '276');
  assert.equal(normalizeCountryCode('SE'), '840');
  assert.equal(normalizeCountryCode('d{3}', '578'), '578');
});

test('import bounds identify row and column truncation', () => {
  assert.deepEqual(importBounds(4999, 99, 5000, 100), {
    lastRow: 4999,
    lastColumn: 99,
    truncatedRows: false,
    truncatedColumns: false,
    sourceRows: 5000,
    sourceColumns: 100,
  });
  const limited = importBounds(7000, 120, 5000, 100);
  assert.equal(limited.lastRow, 4999);
  assert.equal(limited.lastColumn, 99);
  assert.equal(limited.truncatedRows, true);
  assert.equal(limited.truncatedColumns, true);
});

test('wire positions are clamped and untrusted persisted values are removed', () => {
  assert.deepEqual(clampWirePosition({ x: -30, y: 900 }, { width: 1000, height: 600, nodeWidth: 190, nodeHeight: 42 }), { x: 0, y: 558 });
  assert.deepEqual(sanitizeWirePositions({ overview: { a: { x: 10.123, y: 20.456 }, bad: { x: 'no', y: 4 } } }), {
    overview: { a: { x: 10.12, y: 20.46 } },
    underhood: {},
  });
});

test('profitability is incomplete when rates, prices, or freight are missing', () => {
  assert.deepEqual(assessProfitabilityCompleteness({ items: [{ discountedTotal: 10, unitSalesPrice: 20, unitLandedCost: 10 }] }), {
    complete: true,
    missingItemCount: 0,
    reasons: [],
  });
  const incomplete = assessProfitabilityCompleteness({
    warehouseRateMissing: true,
    items: [{ missingFreight: true, discountedTotal: 10, unitSalesPrice: null, unitLandedCost: null }],
  });
  // A part with no stock has no unit price by design, so it is not missing data.
  assert.equal(assessProfitabilityCompleteness({ items: [{ maxQuantity: 0, discountedTotal: 0, unitSalesPrice: null, unitLandedCost: null }] }).complete, true);
  assert.equal(incomplete.complete, false);
  assert.equal(incomplete.missingItemCount, 1);
  assert.equal(incomplete.reasons.length, 2);
});

test('common parts use the maximum quantity found across source sheets', () => {
  assert.deepEqual(summarizeOccurrences([
    { sheetName: 'Kit A', quantity: 2 },
    { sheetName: 'Kit B', quantity: 9 },
    { sheetName: 'Kit A', quantity: 4 },
  ]), { sources: ['Kit A', 'Kit B'], common: true, maxQuantity: 9 });
});

test('freight allocation supports value, quantity, and item-line weighting', () => {
  const items = [{ maxQuantity: 1, discountedTotal: 20 }, { maxQuantity: 3, discountedTotal: 80 }];
  assert.deepEqual(allocationShares(items, 'value'), [0.2, 0.8]);
  assert.deepEqual(allocationShares(items, 'quantity'), [0.25, 0.75]);
  assert.deepEqual(allocationShares(items, 'lines'), [0.5, 0.5]);
});

test('zero warehouse units require zero locations', () => {
  assert.equal(requiredLocations(0, 8), 0);
  assert.equal(requiredLocations(1, 8), 1);
  assert.equal(requiredLocations(17, 8), 3);
});

test('warehouse storage plan uses shelves by default and exact planned pallet counts', () => {
  const shelves = planWarehouseStorage({ inventoryUnits: 100, unitsPerShelf: 20, unitsPerDrawer: 8, unitsPerPallet: 60 });
  assert.equal(shelves.shelfLocations, 5);
  assert.equal(shelves.drawerLocations, 0);
  assert.equal(shelves.pallets, 0);
  assert.equal(shelves.capacityShortfall, 0);

  const mixed = planWarehouseStorage({ inventoryUnits: 100, shelfEnabled: true, palletEnabled: true, plannedPallets: 1, unitsPerShelf: 20, unitsPerDrawer: 8, unitsPerPallet: 60 });
  assert.equal(mixed.pallets, 1);
  assert.equal(mixed.shelfLocations, 2);
  assert.equal(mixed.capacity, 100);

  const short = planWarehouseStorage({ inventoryUnits: 100, shelfEnabled: false, palletEnabled: true, plannedPallets: 1, unitsPerPallet: 60 });
  assert.equal(short.pallets, 1);
  assert.equal(short.capacityShortfall, 40);
});

test('scenario projection separates sales, landed purchase, inbound freight, and warehouse changes', () => {
  const model = {
    rows: [{ sales: 10, item: { maxQuantity: 10, discountedTotal: 500 } }],
    revenue: 1000,
    cogs: 600,
    storage: 50,
    outtake: 20,
    outbound: 30,
    fixed: 100,
    investment: 600,
  };
  const scenario = projectScenario(model, [{ discountedTotal: 500 }], { salesDelta: 20, purchaseDelta: 10, inboundDelta: 50, warehouseDelta: 20 });
  assert.equal(scenario.revenue, 1200);
  assert.equal(scenario.cogs, 840);
  assert.equal(scenario.storage, 60);
  assert.equal(scenario.variableLogistics, 72);
  assert.equal(scenario.fixed, 120);
  assert.equal(scenario.net, 108);
  assert.equal(scenario.investment, 700);
});

test('customs treatment depends on the origin and the warehouse country', () => {
  // Sweden as the warehouse: Bulgaria is the EU internal market, the USA and Switzerland are real imports.
  assert.equal(customsTreatment('100', '752'), 'intraeu');
  assert.equal(customsTreatment('840', '752'), 'import');
  assert.equal(customsTreatment('756', '752'), 'import');
  assert.equal(customsTreatment('752', '752'), 'domestic');
  // Norway is outside the EU customs union: every foreign origin is an import, including EU countries.
  assert.equal(customsTreatment('100', '578'), 'import');
  assert.equal(customsTreatment('756', '578'), 'import');
  assert.equal(customsTreatment('578', '578'), 'domestic');
  // Unknown or malformed codes are treated conservatively as imports.
  assert.equal(customsTreatment('SE', '752'), 'import');
  assert.equal(customsTreatment('', ''), 'import');
  assert.ok(EU_MEMBERS.includes('752') && !EU_MEMBERS.includes('578') && !EU_MEMBERS.includes('756'));
});

test('VAT is off by default and, once a rate is entered, follows the sales country', () => {
  const treatment = { 578: 'none', 752: 'price' };
  // Nothing is added while the project rate is 0, wherever the goods are sold.
  assert.equal(salesVatRate(0, treatment, '578'), 0);
  assert.equal(salesVatRate(0, treatment, '752'), 0);
  // With 25 %: sales to Norway carry no VAT (USA -> Norway, Sweden -> Norway) ...
  assert.equal(salesVatRate(0.25, treatment, '578'), 0);
  // ... and sales to Sweden do (Sweden -> Sweden, Norway -> Sweden, Switzerland -> Sweden).
  assert.equal(salesVatRate(0.25, treatment, '752'), 0.25);
  // The per-country setting can be changed, and unknown countries or bad rates add nothing.
  assert.equal(salesVatRate(0.25, { 578: 'price' }, '578'), 0.25);
  assert.equal(salesVatRate(0.25, treatment, '999'), 0);
  assert.equal(salesVatRate('x', treatment, '752'), 0);
  assert.equal(salesVatRate(-0.1, treatment, '752'), 0);
});

test('the map projection fills its panel, keeps north up, and places the route endpoints inside it', () => {
  const panel = fitProjection([-125, 26, -6, 74], 600, 480);
  const colorado = panel.project([-107.9, 38.5]);
  const norwayCoast = panel.project([-8, 60]);
  assert.ok(colorado[0] > 0 && colorado[0] < 600 && colorado[1] > 0 && colorado[1] < 480);
  assert.ok(norwayCoast[0] > colorado[0]);            // east is to the right
  assert.ok(norwayCoast[1] < colorado[1]);            // 60N is above 38.5N
  // The bounds are expanded so the projection covers exactly the panel size.
  const [west, south, east, north] = panel.bounds;
  const corner = panel.project([west, north]);
  const opposite = panel.project([east, south]);
  assert.ok(Math.abs(corner[0]) < 1e-6 && Math.abs(corner[1]) < 1e-6);
  assert.ok(Math.abs(opposite[0] - 600) < 1e-6 && Math.abs(opposite[1] - 480) < 1e-6);
});

test('a route line can be drawn up to a box and continue on the far side', () => {
  const line = (t) => [t * 1000, 200];                 // straight, left to right
  const box = rectAround([500, 200], { w: 200, h: 300 });
  const around = curveAroundRect(line, box);
  assert.ok(around.enter > 0.39 && around.enter < 0.4);
  assert.ok(around.exit > 0.6 && around.exit < 0.61);
  assert.equal(curveAroundRect(line, rectAround([500, 600], { w: 200, h: 100 })), null);
  assert.ok(rectsOverlap(box, rectAround([630, 200], { w: 100, h: 100 })));
  assert.ok(!rectsOverlap(box, rectAround([900, 200], { w: 100, h: 100 })));
  assert.ok(rectContains(box, [500, 200]) && !rectContains(box, [100, 200]));
});

test('an origin outside a panel is shown on the panel border in its direction', () => {
  const panel = { x: 0, y: 0, w: 400, h: 300 };
  const [x, y] = rayToRect([200, 150], [200, 900], panel, 10);     // straight south
  assert.equal(Math.round(x), 200);
  assert.equal(Math.round(y), 290);
  const [ex, ey] = rayToRect([200, 150], [900, 150], panel);       // straight east
  assert.equal(Math.round(ex), 400);
  assert.equal(Math.round(ey), 150);
});

test('route arcs start and end on their points and bow towards the top of the map', () => {
  const across = quadraticBow([0, 300], [600, 300], 0.2);
  assert.deepEqual(across.at(0), [0, 300]);
  assert.deepEqual(across.at(1), [600, 300]);
  assert.ok(across.at(0.5)[1] < 300);                          // west to east bows north
  assert.ok(Math.abs(across.angle(0.5)) < 1e-6);               // level at the top of the arch
  const north = quadraticBow([100, 500], [100, 100], 0.1);     // a vertical line bows towards the west
  assert.ok(north.at(0.5)[0] < 100);
  assert.ok(north.angle(0.5) < -80 && north.angle(0.5) > -100);  // pointing up the map
});

test('a box is placed on its anchor when clear, and otherwise as near as possible without covering anything', () => {
  const bounds = { x: 0, y: 0, w: 600, h: 400 };
  const size = { w: 200, h: 100 };
  const free = placeBox(size, [300, 200], { bounds });
  assert.ok(free.ok && free.x === 200 && free.y === 150);
  const blocker = rectAround([300, 200], { w: 100, h: 100 });
  const moved = placeBox(size, [300, 200], { bounds, avoid: [blocker] });
  assert.ok(moved.ok);
  assert.ok(!rectsOverlap(moved, blocker));
  assert.ok(Math.hypot(moved.x + 100 - 300, moved.y + 50 - 200) < 130);      // still close to where it was wanted
  const noLeft = placeBox(size, [300, 200], { bounds, blocked: (rect) => rect.x < 250 });   // a region the box must stay out of
  assert.ok(noLeft.ok && noLeft.x >= 250);
  const edge = placeBox(size, [590, 390], { bounds });                         // wanted outside the panel: pulled inside
  assert.ok(edge.ok && edge.x + edge.w <= 600 && edge.y + edge.h <= 400);
  const crowded = placeBox(size, [300, 200], { bounds, avoid: [{ x: 0, y: 0, w: 600, h: 400 }] });
  assert.equal(crowded.ok, false);                                             // nothing is clear, but it stays inside
  assert.ok(crowded.x >= 0 && crowded.y >= 0);
});

test('consolidated sorting orders numbers numerically, text naturally, and keeps blanks last in both directions', () => {
  const items = [
    { key: 'b', part: 'B', total: 80 },
    { key: 'a10', part: 'A10', total: 1200 },
    { key: 'a2', part: 'A2', total: null },
    { key: 'c', part: 'C', total: 300 },
  ];
  const byPart = sortItems(items, { key: 'part', direction: 'asc' }, { keyFor: (item) => item.part });
  // "A2" before "A10" means a numeric collation, not plain string ordering.
  assert.deepEqual(byPart.map((item) => item.key), ['a2', 'a10', 'b', 'c']);

  const byTotalAsc = sortItems(items, { key: 'total', direction: 'asc' }, { keyFor: (item) => item.total });
  assert.deepEqual(byTotalAsc.map((item) => item.key), ['b', 'c', 'a10', 'a2']);

  // 1200 must sort above 80; a string sort would put it first.
  const byTotalDesc = sortItems(items, { key: 'total', direction: 'desc' }, { keyFor: (item) => item.total });
  assert.deepEqual(byTotalDesc.map((item) => item.key), ['a10', 'c', 'b', 'a2']);

  // No sort, an unknown key, or a bad direction leaves the original order alone.
  assert.deepEqual(sortItems(items, null).map((item) => item.key), ['b', 'a10', 'a2', 'c']);
  assert.deepEqual(sortItems(items, { key: 'nope', direction: 'asc' }).map((item) => item.key), ['b', 'a10', 'a2', 'c']);
  assert.deepEqual(sortItems(items, { key: 'part', direction: 'sideways' }).map((item) => item.key), ['b', 'a10', 'a2', 'c']);
});

test('missing inputs are classified by what is actually absent', () => {
  assert.deepEqual(missingInputReasons({ unitPrice: 10, fx: 9.1, missingFreight: false }), []);
  assert.deepEqual(missingInputReasons({ unitPrice: null, fx: 9.1 }), ['price']);
  assert.deepEqual(missingInputReasons({ unitPrice: 10, fx: null }), ['price']);
  assert.deepEqual(missingInputReasons({ unitPrice: 10, fx: 9.1, missingFreight: true }), ['freight']);
  // A row with no price and no freight reports both, without repeating itself.
  assert.deepEqual(missingInputReasons({ unitPrice: null, fx: null, missingFreight: true }), ['price', 'freight']);
  assert.equal(hasMissingInputs({ unitPrice: 1, fx: 1 }), false);
  assert.equal(hasMissingInputs({ unitPrice: null, fx: 1 }), true);
});

test('a catalogue sheet is split into kits; the first article number identifies a part, the second is kit-specific', () => {
  const header = [[], [null, ' 1 / 3'], ['SPARE PART KITS'], [null, null, null, null, null, null, null, null, 'Listprice ', 'Net x '], [null, null, null, 0, null, null, 0, 0, 'EUR', 'EUR']];
  const line = (kit, number, qty, description, article, alternate, list) => [null, kit, number, qty, null, description, article, alternate, list, list * 0.5];
  const rows = [
    ...header,
    line('SPARE PART KIT ZMP 625V', ' ', 4, 'Valve seat', 'A-100', '100.001', 44),
    [null, null, 0, 1, null, 'Oil filter cartridge', 'A-300', 'XXX.300', 115, 57.5],
    [null, null, 0, '2', null, 'Relay', 900001, ' ', 27, 20.25],
    [null, null, 0, 0, null, null, 0, 0],
    ...header.slice(1),
    line('SPARE PART KIT ZMP 710/712V', ' ', 2, 'Valve seat', 'A-200', '200.001', 78.8),
    [null, '(without plunger unit)', null, 1, null, 'Oil filter cartridge', 'A-300', '200.300', 115, 57.5],
    line('SPARE PART HRW 350', 1000337, 1, 'Rubber scraper', 'A-400', '300.001', 54),
  ];
  const catalogue = parseKitCatalogue(rows);
  assert.equal(catalogue.currency, 'EUR');
  assert.deepEqual(catalogue.kits.map((kit) => kit.name), ['ZMP 625V', 'ZMP 710/712V', 'HRW 350']);
  const first = catalogue.kits[0];
  assert.equal(first.lines.length, 3);                       // the zero rows and page headers are not lines
  assert.equal(first.lines[0].article, 'A-100');
  assert.equal(first.lines[0].alternate, '100.001');
  assert.equal(first.lines[0].netPrice, 22);
  assert.equal(first.lines[2].quantity, 2);                  // a quantity typed as text still counts
  assert.equal(first.lines[2].article, '900001');           // the second number may be missing
  assert.equal(first.lines[2].alternate, '');
  assert.deepEqual(catalogue.kits[1].notes, ['(without plunger unit)']);
  assert.equal(catalogue.kits[2].number, '1000337');
  // The same first number in two kits is the same part even though the second number differs.
  assert.equal(first.lines[1].article, catalogue.kits[1].lines[1].article);
  assert.notEqual(first.lines[1].alternate, catalogue.kits[1].lines[1].alternate);
  // An ordinary parts list is not mistaken for a catalogue.
  assert.equal(parseKitCatalogue([['Part', 'Quantity', 'Price'], ['A', 1, 2], ['B', 2, 3]]), null);
});

test('catalogue kits become tidy sheets with safe names', () => {
  assert.equal(catalogueSheetName('KIT 300/600', 'Acme'), 'Acme KIT 300-600');
  assert.equal(catalogueSheetName('A very long kit name that cannot fit', 'Acme').length <= 31, true);
  const rows = catalogueSheetRows({ title: 'SPARE PART KIT X', number: '1000333', notes: [], lines: [{ quantity: 2, description: 'O-Ring', article: 'D-1', alternate: '100.022', netPrice: 1.5000000001, listPrice: 2 }] }, { manufacturer: 'Acme', currency: 'EUR' });
  assert.deepEqual(rows[3], ['Article no.', 'Alt. article no.', 'Description', 'Quantity', 'Unit price (net)', 'Currency', 'List price (info)']);
  assert.deepEqual(rows[4], ['D-1', '100.022', 'O-Ring', 2, 1.5, 'EUR', 2]);
});
test('a sales history is read in both layouts and added up per article', () => {
  const split = parseSalesHistory([
    ['Item No.', 'Description', 'Quantity'],
    ['Total', '', 12],
    ['A-1', 'Valve seat', '4.00'],
    ['a-1 ', 'Valve seat (old name)', 5],
    ['B-2', 'Seal', '1 000,5'],
    ['', 'no number', 3],
  ]);
  assert.equal(split.items.length, 2);
  assert.deepEqual(split.items.map((entry) => [entry.itemNo, entry.quantity]), [['B-2', 1000.5], ['A-1', 9]]);
  assert.equal(split.items[1].description, 'Valve seat (old name)');
  assert.equal(split.reportedTotal, 12);

  const combined = parseSalesHistory([['Item No. / Description', 'Quantity '], ['Total', 7], ['2513-AG-11 - Valve seat ZMP 700', 5], ['D-1', 2]]);
  assert.deepEqual(combined.items.map((entry) => [entry.itemNo, entry.description, entry.quantity]), [['2513-AG-11', 'Valve seat ZMP 700', 5], ['D-1', '', 2]]);

  // Price lists and empty sheets are not histories.
  assert.equal(parseSalesHistory([['Item', 'Quantity', 'Net price'], ['A', 1, 2]]), null);
  assert.equal(parseSalesHistory([['Item No.', 'Quantity'], ['List price', 'Net price'], ['A', 1]]), null);
  assert.equal(parseSalesHistory([['Item No.', 'Quantity']]), null);
  assert.equal(parseSalesHistory([['Hello']]), null);
});

test('popularity joins kit parts to every market and finds what sells in both', () => {
  const parts = [
    { key: 'a-1', part: 'A-1', altParts: ['100.001'], description: 'Valve seat', kits: ['Kit 1', 'Kit 2'] },
    { key: 'b-2', part: 'B-2', altParts: [], description: 'Seal', kits: ['Kit 1'] },
    { key: 'c-3', part: 'C-3', altParts: ['300.003'], description: 'Never sold', kits: ['Kit 2'] },
  ];
  const markets = [
    { name: 'Norway', items: [{ key: 'A-1', itemNo: 'A-1', description: 'x', quantity: 10 }, { key: 'Z-9', itemNo: 'Z-9', description: 'Hose', quantity: 7 }] },
    { key: 'unused', name: 'Sweden', items: [{ key: '100.001', itemNo: '100.001', description: 'x', quantity: 5 }, { key: 'B-2', itemNo: 'B-2', description: 'x', quantity: 30 }, { key: 'Z-9', itemNo: 'Z-9', description: 'Hose', quantity: 1 }] },
  ];
  const result = buildPopularity({ parts, markets });
  const byPart = Object.fromEntries(result.rows.map((row) => [row.part, row]));
  // A part is found by its alternative number too.
  assert.deepEqual(byPart['A-1'].sold, [10, 5]);
  assert.deepEqual(byPart['B-2'].sold, [0, 30]);
  assert.equal(byPart['C-3'].total, 0);
  assert.equal(byPart['C-3'].rank, null);
  // Articles that are in no kit are kept and marked.
  assert.equal(byPart['Z-9'].inKit, false);
  assert.deepEqual(byPart['Z-9'].sold, [7, 1]);
  // Popular in both: only articles sold in every market, ranked by the weaker market.
  const both = result.rows.filter((row) => row.inAll).sort((a, b) => a.bothRank - b.bothRank).map((row) => row.part);
  assert.deepEqual(both, ['A-1', 'Z-9']);
  assert.equal(byPart['B-2'].inAll, false);
  assert.equal(byPart['B-2'].bothRank, null);
  // Ranks: total first, then per market.
  assert.deepEqual(result.rows.map((row) => row.part).slice(0, 4), ['B-2', 'A-1', 'Z-9', 'C-3']);
  assert.equal(byPart['A-1'].rank, 2);
  assert.deepEqual(byPart['A-1'].marketRanks, [1, 2]);
  // Kits add up their parts; a part in two kits counts for both, and nothing sold is not ranked.
  const kit1 = result.kits.find((kit) => kit.name === 'Kit 1');
  const kit2 = result.kits.find((kit) => kit.name === 'Kit 2');
  assert.deepEqual(kit1.sold, [10, 35]);
  assert.equal(kit1.soldParts, 2);
  assert.equal(kit1.top.part, 'B-2');
  assert.deepEqual(kit2.sold, [10, 5]);
  assert.equal(kit2.soldParts, 1);
  assert.equal(result.kits[0].name, 'Kit 1');
  // Coverage shows how much of each market's volume the kits explain.
  assert.deepEqual(result.coverage.map((entry) => [entry.name, entry.units, entry.matchedUnits]), [['Norway', 17, 10], ['Sweden', 36, 35]]);
});

test('suggested stock covers the demand of the cover period with a safety margin and respects a budget', () => {
  const part = (key, total, kits = ['K']) => ({ key, part: key.toUpperCase(), altParts: [], description: '', kits, inKit: true, total, sold: [total] });
  const rows = [part('a', 500), part('b', 100), part('c', 5), part('d', 0), part('e', 0)];
  const members = (names) => rows.filter((row) => names.includes(row.key));
  const kits = [{ name: 'K', parts: 4, soldParts: 3, members: members(['a', 'b', 'c', 'd']) }];
  const plan = suggestStock({ rows, kits, years: 5, monthsCover: 6, serviceLevel: 0.95, unitCost: () => 10, current: () => 2 });
  const by = Object.fromEntries(plan.lines.map((line) => [line.key, line]));
  // 500 pieces in 5 years = 100 a year; six months = 50, plus 1.645 * sqrt(50) ≈ 11.6 → 62.
  assert.equal(by.a.target, 62);
  assert.equal(by.a.cls, 'A');
  assert.equal(by.c.target, 2);                       // 0.5 expected + margin rounds up; a slow seller never gets less than one
  assert.equal(by.c.cls, 'C');
  assert.equal(by.d.target, 1);                       // unsold, but in a kit where most parts sell
  assert.equal(by.d.reason, 'kit');
  assert.equal(by.e.target, 0);                       // unsold and in no kit
  assert.equal(plan.value, (62 + by.b.target + 2 + 1) * 10);
  assert.equal(plan.valueNow, 5 * 2 * 10);
  const without = suggestStock({ rows, kits, years: 5, completeKits: false, unitCost: () => 10 });
  assert.equal(without.lines.find((line) => line.key === 'd').target, 0);
  const tight = suggestStock({ rows, kits, years: 5, budget: 300, unitCost: () => 10 });
  assert.ok(tight.value <= 300);
  assert.equal(tight.overBudget, false);
  assert.ok(tight.lines.find((line) => line.key === 'a').target >= 1);
  assert.ok(tight.lines.find((line) => line.key === 'a').target < 62);
  // A budget that cannot be met even with nothing stocked is reported.
  assert.equal(suggestStock({ rows, kits, years: 5, budget: 0.5, unitCost: () => 10 }).value, 0);
});

test('one article per bin needs whole bins per article, several per bin only the pooled space', () => {
  const units = [3, 3, 3, 12];                 // four articles, 21 pieces, 5 pieces fit in a bin
  assert.equal(binsForUnits(12, 5, true), 3);
  assert.equal(binsForUnits(12, 5, false), 2.4);
  assert.equal(binsForUnits(0, 5, true), 0);
  const base = { inventoryUnits: 21, itemUnits: units, unitsPerBin: 5, unitsPerShelf: 20, binsPerLocation: 4 };
  const shared = planStorage({ ...base, onePerBin: false });
  assert.equal(shared.shelfLocations, 2);       // 21 pieces / 20 per shelf location
  assert.equal(shared.onePerBin, false);
  const single = planStorage({ ...base, onePerBin: true });
  assert.equal(single.shelfBins, 1 + 1 + 1 + 3);  // 3+3+3 pieces need a bin each, 12 pieces three bins
  assert.equal(single.shelfLocations, 2);       // 6 bins / 4 per location
  assert.equal(single.binsWasted, 6 * 5 - 21);
  // Bins stacked high and deep give a shelf location more bins, so fewer locations are needed.
  const deep = planStorage({ ...base, onePerBin: true, binsPerLocation: 12, unitsPerShelf: 60 });
  assert.equal(deep.shelfLocations, 1);
  // Many small articles waste space when they cannot share a bin.
  const many = planStorage({ inventoryUnits: 40, itemUnits: Array(20).fill(2), unitsPerBin: 5, unitsPerShelf: 20, binsPerLocation: 4, onePerBin: true });
  assert.equal(many.shelfBins, 20);
  assert.equal(many.shelfLocations, 5);
  assert.equal(planStorage({ inventoryUnits: 40, unitsPerBin: 5, unitsPerShelf: 20, binsPerLocation: 4 }).shelfLocations, 2);
});

test('with minimal lists a part is common only when two lists need at least one piece of it', () => {
  const rows = [{ sheetName: 'A', quantity: 2 }, { sheetName: 'B', quantity: 0 }];
  assert.equal(summarizeOccurrences(rows).common, true);
  assert.equal(summarizeOccurrences(rows, { requireStock: true }).common, false);
  assert.equal(summarizeOccurrences([...rows, { sheetName: 'C', quantity: 1 }], { requireStock: true }).common, true);
  assert.deepEqual(summarizeOccurrences(rows, { requireStock: true }).sources, ['A', 'B']);
});

test('the bin what-if compares one article per bin with sharing, finds the best bin size and the room left in bins', () => {
  const items = [{ key: 'a', part: 'A', units: 12 }, { key: 'b', part: 'B', units: 3 }, { key: 'c', part: 'C', units: 3 }, { key: 'd', part: 'D', units: 10 }, { key: 'z', part: 'Z', units: 0 }];
  const result = analyzeBins({ items, unitsPerBin: 5, binsPerLocation: 4, candidates: [3, 10, 28] });
  assert.equal(result.articles, 4);                                   // no stock, no bin
  assert.equal(result.totalUnits, 28);
  assert.equal(result.shared.bins, 6);                                // 28 / 5 rounded up
  assert.equal(result.single.bins, 3 + 1 + 1 + 2);                    // A needs 3, B and C one each, D two
  assert.equal(result.extraBins, 1);
  assert.equal(result.shared.locations, 2);                           // 6 bins / 4 per location
  assert.equal(result.single.locations, 2);                           // 7 bins / 4 per location
  assert.equal(result.single.room, 35 - 28);                          // room left in the bins in use
  assert.equal(result.list[0].part, 'A');
  assert.equal(result.list[0].room, 3);
  // Bigger bins mean fewer bins: 28 items per bin is one bin per article.
  const big = result.sweep.find((entry) => entry.unitsPerBin === 28);
  assert.equal(big.bins, 4);
  assert.equal(result.best.locations, 1);
  assert.ok(result.best.unitsPerBin >= 10);
  // Stock can be filled up to whole bins.
  const rows = [{ key: 'a', part: 'A', altParts: [], description: '', kits: ['K'], inKit: true, total: 500, sold: [500] }];
  const plain = suggestStock({ rows, kits: [], years: 5, unitCost: () => 1 }).lines[0].target;
  const filled = suggestStock({ rows, kits: [], years: 5, binSize: 25, unitCost: () => 1 }).lines[0].target;
  assert.equal(filled % 25, 0);
  assert.ok(filled >= plain && filled - plain < 25);
});
