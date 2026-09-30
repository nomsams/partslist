const test = require('node:test');
const assert = require('node:assert/strict');
const {
  assessProfitabilityCompleteness,
  allocationShares,
  clampWirePosition,
  importBounds,
  normalizeCountryCode,
  planWarehouseStorage,
  projectScenario,
  requiredLocations,
  sanitizeWirePositions,
  summarizeOccurrences,
} = require('../partslist-core.js');

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
