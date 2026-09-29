(function exposePartsListCore(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.PartsListCore = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  'use strict';

  function normalizeCountryCode(value, fallback = '840') {
    const candidate = String(value ?? '');
    return /^\d{3}$/.test(candidate) ? candidate : fallback;
  }

  function importBounds(endRow, endColumn, maxRows, maxColumns) {
    const rowLimit = Math.max(1, Math.trunc(maxRows));
    const columnLimit = Math.max(1, Math.trunc(maxColumns));
    return {
      lastRow: Math.min(Math.max(0, endRow), rowLimit - 1),
      lastColumn: Math.min(Math.max(0, endColumn), columnLimit - 1),
      truncatedRows: endRow >= rowLimit,
      truncatedColumns: endColumn >= columnLimit,
      sourceRows: Math.max(0, endRow) + 1,
      sourceColumns: Math.max(0, endColumn) + 1,
    };
  }

  function clampWirePosition(position, bounds) {
    const x = Number(position?.x);
    const y = Number(position?.y);
    const maxX = Math.max(0, Number(bounds?.width || 0) - Number(bounds?.nodeWidth || 0));
    const maxY = Math.max(0, Number(bounds?.height || 0) - Number(bounds?.nodeHeight || 0));
    return {
      x: Math.min(maxX, Math.max(0, Number.isFinite(x) ? x : 0)),
      y: Math.min(maxY, Math.max(0, Number.isFinite(y) ? y : 0)),
    };
  }

  function sanitizeWirePositions(value) {
    const result = { overview: {}, underhood: {} };
    ['overview', 'underhood'].forEach((mode) => {
      const positions = value?.[mode];
      if (!positions || typeof positions !== 'object' || Array.isArray(positions)) return;
      Object.entries(positions).slice(0, 500).forEach(([id, position]) => {
        const x = Number(position?.x);
        const y = Number(position?.y);
        if (!id || id.length > 160 || !Number.isFinite(x) || !Number.isFinite(y)) return;
        result[mode][id] = { x: Math.round(x * 100) / 100, y: Math.round(y * 100) / 100 };
      });
    });
    return result;
  }

  function assessProfitabilityCompleteness({ warehouseRateMissing = false, items = [] } = {}) {
    const missingItems = items.filter((item) => (
      item?.missingFreight
      || item?.discountedTotal === null
      || item?.unitSalesPrice === null
      || item?.unitLandedCost === null
    ));
    const reasons = [];
    if (warehouseRateMissing) reasons.push('NOK conversion rate');
    if (missingItems.length) reasons.push(`${missingItems.length} part${missingItems.length === 1 ? '' : 's'} with missing price or freight data`);
    return {
      complete: reasons.length === 0,
      missingItemCount: missingItems.length,
      reasons,
    };
  }

  function summarizeOccurrences(occurrences = []) {
    const sources = [...new Set(occurrences.map((item) => item.sheetName).filter(Boolean))];
    const quantities = occurrences.map((item) => Number(item.quantity)).filter(Number.isFinite);
    return {
      sources,
      common: sources.length > 1,
      maxQuantity: quantities.length ? Math.max(...quantities) : 0,
    };
  }

  function allocationShares(items = [], mode = 'value') {
    const weight = (item) => {
      if (mode === 'quantity') return Math.max(0, Number(item?.maxQuantity) || 0);
      if (mode === 'lines') return Number(item?.maxQuantity) > 0 ? 1 : 0;
      return Math.max(0, Number(item?.discountedTotal) || 0);
    };
    const weights = items.map(weight);
    const total = weights.reduce((sum, value) => sum + value, 0);
    return weights.map((value) => total > 0 ? value / total : 0);
  }

  function requiredLocations(units, capacity) {
    const amount = Math.max(0, Number(units) || 0);
    const perLocation = Math.max(1, Number(capacity) || 1);
    return amount > 0 ? Math.ceil(amount / perLocation) : 0;
  }

  function projectScenario(model, inventoryItems = [], changes = {}) {
    const factor = (delta) => Math.max(0, 1 + (Number(delta) || 0) / 100);
    const salesFactor = factor(changes.salesDelta);
    const purchaseFactor = factor(changes.purchaseDelta);
    const inboundFactor = factor(changes.inboundDelta);
    const warehouseFactor = factor(changes.warehouseDelta);
    const purchaseCogs = (model.rows || []).reduce((sum, row) => {
      const stock = Math.max(0, Number(row.item?.maxQuantity) || 0);
      const purchase = row.item?.discountedTotal;
      const unitPurchase = stock > 0 && purchase !== null && Number.isFinite(purchase) ? purchase / stock : 0;
      return sum + (Number(row.sales) || 0) * unitPurchase;
    }, 0);
    const inboundCogs = Math.max(0, (Number(model.cogs) || 0) - purchaseCogs);
    const revenue = (Number(model.revenue) || 0) * salesFactor;
    const cogs = purchaseCogs * salesFactor * purchaseFactor + inboundCogs * salesFactor * inboundFactor;
    const storage = (Number(model.storage) || 0) * warehouseFactor;
    const variableLogistics = ((Number(model.outtake) || 0) + (Number(model.outbound) || 0)) * salesFactor * warehouseFactor;
    const fixed = (Number(model.fixed) || 0) * warehouseFactor;
    const gross = revenue - cogs;
    const net = gross - storage - variableLogistics - fixed;
    const purchaseInvestment = inventoryItems.reduce((sum, item) => sum + (Number(item?.discountedTotal) || 0), 0);
    const inboundInvestment = Math.max(0, (Number(model.investment) || 0) - purchaseInvestment);
    const investment = purchaseInvestment * purchaseFactor + inboundInvestment * inboundFactor;
    const monthlyNet = net / 12;
    return {
      revenue,
      cogs,
      gross,
      storage,
      variableLogistics,
      fixed,
      net,
      investment,
      margin: revenue > 0 ? net / revenue : null,
      paybackMonths: investment > 0 && monthlyNet > 0 ? investment / monthlyNet : null,
    };
  }

  return {
    assessProfitabilityCompleteness,
    allocationShares,
    clampWirePosition,
    importBounds,
    normalizeCountryCode,
    projectScenario,
    requiredLocations,
    sanitizeWirePositions,
    summarizeOccurrences,
  };
}));
