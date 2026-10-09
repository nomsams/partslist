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

  // ISO 3166-1 numeric codes of the EU member states.
  const EU_MEMBERS = Object.freeze(['040', '056', '100', '191', '196', '203', '208', '233', '246', '250', '276', '300', '348', '372', '380', '428', '440', '442', '470', '528', '616', '620', '642', '703', '705', '724', '752']);

  // How goods travelling from an origin country are treated on arrival in the warehouse country:
  // 'domestic' (no border), 'intraeu' (EU internal market: no duty, no import declaration; acquisition VAT is
  // reverse-charged), or 'import' (customs clearance, duty and import VAT).
  function customsTreatment(origin, destination) {
    const from = normalizeCountryCode(origin, '');
    const to = normalizeCountryCode(destination, '');
    if (!from || !to) return 'import';
    if (from === to) return 'domestic';
    if (EU_MEMBERS.includes(from) && EU_MEMBERS.includes(to)) return 'intraeu';
    return 'import';
  }

  // VAT is a sales-price add-on, applied after the margin, and only for sales countries set to "price".
  // With the project rate at 0 (the default) nothing is added anywhere.
  function salesVatRate(rate, treatmentByCountry, customerCountry) {
    const value = Number(rate);
    if (!Number.isFinite(value) || value <= 0) return 0;
    return treatmentByCountry?.[normalizeCountryCode(customerCountry, '')] === 'price' ? value : 0;
  }

  // ---- Kit catalogues (several kits in one price-list sheet) ----------------------------------------
  // A catalogue sheet lists many kits one after another. Each kit starts on a row whose column B reads
  // "SPARE PART KIT <name>" and runs until the next such row; page headers repeat in between. A line holds a
  // quantity, a description, the manufacturer's own article number, a second (kit-specific) number, and a list
  // and a net price. The price columns are found from their header text, the other columns relative to them.
  const CATALOGUE_KIT_START = /^spare\s+part(?:\s+kits?)?\s+(.+)$/i;

  function cellText(value) {
    return value === null || value === undefined ? '' : String(value).trim();
  }

  function cellNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    const text = cellText(value).replace(',', '.');
    return /^-?\d+(?:\.\d+)?$/.test(text) ? Number(text) : null;
  }

  // A number such as "XXX.539" is a placeholder: the real number differs from kit to kit and is not filled in.
  function usableArticleNumber(text) {
    return text !== '' && text !== '0' && !/^x{2,}/i.test(text);
  }

  function parseKitCatalogue(rows) {
    if (!Array.isArray(rows)) return null;
    let listColumn = -1;
    let netColumn = -1;
    let currency = '';
    for (let r = 0; r < Math.min(rows.length, 40) && netColumn < 0; r += 1) {
      const row = rows[r] || [];
      const list = row.findIndex((cell) => /^list\s*price/i.test(cellText(cell)));
      const net = row.findIndex((cell) => /^net\b/i.test(cellText(cell)));
      if (list < 0 || net < 0) continue;
      listColumn = list;
      netColumn = net;
      const codes = [rows[r + 1], rows[r + 2]].map((below) => cellText(below?.[net]).toUpperCase());
      currency = codes.find((code) => /^[A-Z]{3}$/.test(code)) || '';
    }
    if (netColumn < 0 || listColumn < 3) return null;
    const alternateColumn = listColumn - 1;
    const articleColumn = listColumn - 2;
    const descriptionColumn = listColumn - 3;

    const kits = [];
    let kit = null;
    rows.forEach((row) => {
      if (!row) return;
      const heading = cellText(row[1]).match(CATALOGUE_KIT_START);
      if (heading && !/^kits?$/i.test(heading[1].trim())) {
        const number = cellNumber(row[2]);
        kit = { name: heading[1].trim(), title: cellText(row[1]), number: number !== null && number > 1000 ? String(number) : '', notes: [], lines: [] };
        kits.push(kit);
      }
      if (!kit) return;
      // Other text in column B belongs to the kit, for example "(without plunger unit)"; "1 / 3" is a page number.
      const note = cellText(row[1]);
      if (note && !heading && !/^\d+\s*\/\s*\d+$/.test(note) && !kit.notes.includes(note)) kit.notes.push(note);
      const net = cellNumber(row[netColumn]);
      const list = cellNumber(row[listColumn]);
      const description = cellText(row[descriptionColumn]);
      if ((net === null && list === null) || !description || cellNumber(description) !== null) return;
      let quantity = null;
      for (let column = descriptionColumn - 1; column >= 0 && quantity === null; column -= 1) {
        const value = cellNumber(row[column]);
        if (value !== null && value > 0 && value < 10000) quantity = value;
        else if (cellText(row[column]) !== '' && value === null) break;
      }
      if (quantity === null) return;
      const primary = cellText(row[articleColumn]);
      const secondary = cellText(row[alternateColumn]);
      const article = usableArticleNumber(primary) ? primary : usableArticleNumber(secondary) ? secondary : description;
      kit.lines.push({
        quantity,
        description,
        article,
        alternate: secondary === '0' ? '' : secondary,
        listPrice: list,
        netPrice: net !== null ? net : list,
      });
    });
    const usable = kits.filter((candidate) => candidate.lines.length);
    return usable.length >= 2 ? { currency, kits: usable } : null;
  }

  function catalogueSheetName(kitName, manufacturer) {
    const clean = (text) => String(text).replace(/\s*[\\/]\s*/g, '-').replace(/[?*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim();
    return `${clean(manufacturer)} ${clean(kitName)}`.trim().slice(0, 31);
  }

  // One tidy sheet per kit in a layout the column detection reads without help.
  function catalogueSheetRows(kit, { manufacturer = '', currency = '' } = {}) {
    const round = (value) => (value === null ? null : Math.round(value * 10000) / 10000);
    return [
      [[manufacturer, kit.title].filter(Boolean).join(' · ')],
      [[kit.number ? `Kit no. ${kit.number}` : '', ...kit.notes, currency ? `Prices in ${currency} (net)` : ''].filter(Boolean).join(' · ')],
      [],
      ['Article no.', 'Alt. article no.', 'Description', 'Quantity', 'Unit price (net)', 'Currency', 'List price (info)'],
      ...kit.lines.map((line) => [line.article, line.alternate, line.description, line.quantity, round(line.netPrice), currency, round(line.listPrice)]),
    ];
  }

  // ---- Route map geometry ----------------------------------------------------------------------------
  // Equirectangular projection that fills exactly width x height around the centre of the given
  // [lonMin, latMin, lonMax, latMax] bounds. project([lon, lat]) returns [x, y] with north up.
  function fitProjection(bounds, width, height) {
    const [lonMin, latMin, lonMax, latMax] = bounds;
    const centreLon = (lonMin + lonMax) / 2;
    const centreLat = (latMin + latMax) / 2;
    const cos = Math.cos(centreLat * Math.PI / 180);
    const scale = Math.min(width / ((lonMax - lonMin) * cos), height / (latMax - latMin));
    const west = centreLon - width / (2 * scale * cos);
    const north = centreLat + height / (2 * scale);
    return {
      scale,
      bounds: [west, centreLat - height / (2 * scale), centreLon + width / (2 * scale * cos), north],
      project: ([lon, lat]) => [(lon - west) * cos * scale, (north - lat) * scale],
    };
  }

  function rectAround([x, y], { w, h }) {
    return { x: x - w / 2, y: y - h / 2, w, h };
  }

  function rectsOverlap(a, b, margin = 0) {
    return a.x < b.x + b.w + margin && b.x < a.x + a.w + margin && a.y < b.y + b.h + margin && b.y < a.y + a.h + margin;
  }

  function rectContains(rect, [x, y], margin = 0) {
    return x >= rect.x - margin && x <= rect.x + rect.w + margin && y >= rect.y - margin && y <= rect.y + rect.h + margin;
  }

  // For a curve given as pointAt(t), t in [0, 1]: the last t before it enters the rectangle and the first t
  // after it leaves, so a line can be drawn up to a box and continue on the far side. null if it never enters.
  function curveAroundRect(pointAt, rect, steps = 240) {
    let first = -1;
    let last = -1;
    for (let index = 0; index <= steps; index += 1) {
      if (!rectContains(rect, pointAt(index / steps))) continue;
      if (first < 0) first = index;
      last = index;
    }
    if (first < 0) return null;
    return { enter: Math.max(0, (first - 1) / steps), exit: Math.min(1, (last + 1) / steps) };
  }

  // A quadratic curve from p to q that bows sideways by `bow` times its length, towards the top of the map
  // (towards the west for a vertical line), so routes read as arcs. at(t) is the point, angle(t) the heading in degrees.
  function quadraticBow([x1, y1], [x2, y2], bow = 0.16) {
    const dx = x2 - x1;
    const dy = y2 - y1;
    const length = Math.hypot(dx, dy) || 1;
    let nx = dy / length;
    let ny = -dx / length;
    if (ny > 1e-9 || (Math.abs(ny) <= 1e-9 && nx > 0)) {
      nx = -nx;
      ny = -ny;
    }
    const cx = (x1 + x2) / 2 + nx * length * bow;
    const cy = (y1 + y2) / 2 + ny * length * bow;
    return {
      control: [cx, cy],
      at: (t) => {
        const u = 1 - t;
        return [u * u * x1 + 2 * u * t * cx + t * t * x2, u * u * y1 + 2 * u * t * cy + t * t * y2];
      },
      angle: (t) => Math.atan2(2 * (1 - t) * (cy - y1) + 2 * t * (y2 - cy), 2 * (1 - t) * (cx - x1) + 2 * t * (x2 - cx)) * 180 / Math.PI,
    };
  }

  // Finds a spot for a box of the given size as close to `anchor` as possible that stays inside `bounds` and
  // clear of every rectangle in `avoid` (and of anything `blocked(rect)` rejects, such as land), searching outwards in square rings. ok is false when nothing is clear.
  function placeBox(size, anchor, { bounds, avoid = [], blocked = null, step = 12, gap = 4 }) {
    const fits = (rect) => rect.x >= bounds.x && rect.y >= bounds.y
      && rect.x + rect.w <= bounds.x + bounds.w && rect.y + rect.h <= bounds.y + bounds.h
      && !avoid.some((other) => rectsOverlap(rect, other, gap))
      && !(blocked && blocked(rect));
    const maxRing = Math.ceil(Math.max(bounds.w, bounds.h) / step);
    for (let ring = 0; ring <= maxRing; ring += 1) {
      let best = null;
      for (let i = -ring; i <= ring; i += 1) {
        for (let j = -ring; j <= ring; j += 1) {
          if (Math.max(Math.abs(i), Math.abs(j)) !== ring) continue;
          const rect = rectAround([anchor[0] + i * step, anchor[1] + j * step], size);
          const distance = Math.hypot(i, j);
          if ((!best || distance < best.distance) && fits(rect)) best = { rect, distance };
        }
      }
      if (best) return { ...best.rect, ok: true };
    }
    const rect = rectAround(anchor, size);
    const clamp = (value, low, high) => Math.max(low, Math.min(value, Math.max(low, high)));
    return { x: clamp(rect.x, bounds.x, bounds.x + bounds.w - size.w), y: clamp(rect.y, bounds.y, bounds.y + bounds.h - size.h), w: size.w, h: size.h, ok: false };
  }

  // The point where the ray from a centre inside a rectangle towards a target meets the rectangle's border.
  function rayToRect([cx, cy], [tx, ty], rect, inset = 0) {
    const dx = tx - cx;
    const dy = ty - cy;
    const left = rect.x + inset;
    const right = rect.x + rect.w - inset;
    const top = rect.y + inset;
    const bottom = rect.y + rect.h - inset;
    const hitX = dx > 0 ? (right - cx) / dx : dx < 0 ? (left - cx) / dx : Infinity;
    const hitY = dy > 0 ? (bottom - cy) / dy : dy < 0 ? (top - cy) / dy : Infinity;
    const t = Math.min(hitX, hitY, 1);
    return [cx + dx * t, cy + dy * t];
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
    // A part with no stock has no unit price or unit cost by design; that is not missing data.
    const missingItems = items.filter((item) => (
      item?.missingFreight
      || item?.discountedTotal === null
      || (!(item?.maxQuantity <= 0) && (item?.unitSalesPrice === null || item?.unitLandedCost === null))
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

  // A part is common when it is in more than one list. With requireStock only lists that need at least one piece of it
  // count, which is how a minimal list works: a part with 0 pieces is listed there but not required.
  function summarizeOccurrences(occurrences = [], { requireStock = false } = {}) {
    const sources = [...new Set(occurrences.map((item) => item.sheetName).filter(Boolean))];
    const quantities = occurrences.map((item) => Number(item.quantity)).filter(Number.isFinite);
    const counted = requireStock
      ? [...new Set(occurrences.filter((item) => Number(item.quantity) > 0).map((item) => item.sheetName).filter(Boolean))]
      : sources;
    return {
      sources,
      common: counted.length > 1,
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

  // Bins one article needs. Alone in its bins, an article fills whole bins (a part-filled bin counts as a bin);
  // sharing bins, it only takes the fraction of a bin that it fills. All items are assumed to be the same size.
  function binsForUnits(units, unitsPerBin, onePerBin = false) {
    const amount = Math.max(0, Number(units) || 0);
    const perBin = Math.max(1, Number(unitsPerBin) || 1);
    if (amount <= 0) return 0;
    return onePerBin ? Math.ceil(amount / perBin - 1e-9) : amount / perBin;
  }

  // options.onePerBin with options.itemUnits (pieces per article) gives every article its own bins; unitsPerShelf
  // is then bins per shelf location x unitsPerBin, and binsPerLocation says how many bins a shelf location holds
  // (bins across x bins stacked high x bins deep).
  // What-if for the shelves: every article in its own bins (one article number per bin, several items inside) against
  // articles sharing bins. All items are assumed to be the same size. items: [{ key, part, description, units, sold }],
  // units = pieces that go on shelves. candidates: other numbers of items per bin to try.
  function analyzeBins({ items = [], unitsPerBin = 1, binsPerLocation = 1, candidates = [] } = {}) {
    const perBin = Math.max(1, Number(unitsPerBin) || 1);
    const perLocation = Math.max(1, Math.round(Number(binsPerLocation) || 1));
    const locationsFor = (bins) => (bins > 0 ? Math.ceil(bins / perLocation) : 0);
    const binsFor = (units, size) => (units > 0 ? Math.ceil(units / size - 1e-9) : 0);
    const list = items.filter((item) => item.units > 0).map((item) => {
      const bins = binsFor(item.units, perBin);
      return { ...item, bins, room: bins * perBin - item.units, fill: item.units / (bins * perBin) };
    }).sort((a, b) => b.bins - a.bins || b.units - a.units);
    const totalUnits = list.reduce((sum, item) => sum + item.units, 0);
    const singleBins = list.reduce((sum, item) => sum + item.bins, 0);
    const sharedBins = binsFor(totalUnits, perBin);
    const sweep = [...new Set([perBin, ...candidates])].filter((size) => size >= 1).sort((a, b) => a - b).map((size) => {
      const bins = list.reduce((sum, item) => sum + binsFor(item.units, size), 0);
      return { unitsPerBin: size, bins, locations: locationsFor(bins) };
    });
    const best = sweep.reduce((winner, entry) => (!winner || entry.locations < winner.locations || (entry.locations === winner.locations && entry.bins < winner.bins) ? entry : winner), null);
    return {
      perBin,
      totalUnits,
      articles: list.length,
      shared: { bins: sharedBins, locations: locationsFor(sharedBins), fill: sharedBins > 0 ? totalUnits / (sharedBins * perBin) : 0 },
      single: { bins: singleBins, locations: locationsFor(singleBins), fill: singleBins > 0 ? totalUnits / (singleBins * perBin) : 0, room: list.reduce((sum, item) => sum + item.room, 0) },
      extraBins: singleBins - sharedBins,
      extraLocations: locationsFor(singleBins) - locationsFor(sharedBins),
      list,
      sweep,
      best,
    };
  }

  function planWarehouseStorage(options = {}) {
    const inventoryUnits = Math.max(0, Number(options.inventoryUnits) || 0);
    const shelfEnabled = options.shelfEnabled !== false;
    const drawerEnabled = options.drawerEnabled === true;
    const palletEnabled = options.palletEnabled === true;
    const unitsPerShelf = Math.max(1, Number(options.unitsPerShelf) || 1);
    const unitsPerDrawer = Math.max(1, Number(options.unitsPerDrawer) || 1);
    const unitsPerPallet = Math.max(1, Number(options.unitsPerPallet) || 1);
    const pallets = palletEnabled ? Math.max(0, Math.round(Number(options.plannedPallets) || 0)) : 0;
    const palletCapacity = pallets * unitsPerPallet;
    const palletUnits = Math.min(inventoryUnits, palletCapacity);
    const remainingUnits = Math.max(0, inventoryUnits - palletCapacity);
    let shelfWeight = shelfEnabled ? 1 : 0;
    let drawerWeight = drawerEnabled ? 1 : 0;
    if (shelfEnabled && drawerEnabled) {
      shelfWeight = Math.max(0, Number(options.shelfShare) || 0);
      drawerWeight = Math.max(0, Number(options.drawerShare) || 0);
      if (shelfWeight + drawerWeight === 0) shelfWeight = drawerWeight = 1;
    }
    const activeWeight = shelfWeight + drawerWeight;
    const shelfUnits = activeWeight ? remainingUnits * shelfWeight / activeWeight : 0;
    const drawerUnits = activeWeight ? remainingUnits * drawerWeight / activeWeight : 0;
    const unitsPerBin = Math.max(1, Number(options.unitsPerBin) || 1);
    const binsPerLocation = Math.max(1, Math.round(Number(options.binsPerLocation) || unitsPerShelf / unitsPerBin));
    const onePerBin = options.onePerBin === true && Array.isArray(options.itemUnits);
    let shelfBins;
    if (onePerBin) {
      // Pallets and drawers take their share of every article; what is left on the shelves is split into whole bins per article.
      const fraction = inventoryUnits > 0 ? shelfUnits / inventoryUnits : 0;
      shelfBins = options.itemUnits.reduce((sum, units) => sum + binsForUnits((Number(units) || 0) * fraction, unitsPerBin, true), 0);
    } else {
      shelfBins = binsForUnits(shelfUnits, unitsPerBin, true);
    }
    const shelfLocations = onePerBin ? Math.ceil(shelfBins / binsPerLocation) : requiredLocations(shelfUnits, unitsPerShelf);
    const drawerLocations = requiredLocations(drawerUnits, unitsPerDrawer);
    const capacity = shelfLocations * unitsPerShelf + drawerLocations * unitsPerDrawer + palletCapacity;
    return {
      shelfEnabled,
      drawerEnabled,
      palletEnabled,
      shelfUnits,
      drawerUnits,
      palletUnits,
      shelfShare: inventoryUnits ? shelfUnits / inventoryUnits : 0,
      drawerShare: inventoryUnits ? drawerUnits / inventoryUnits : 0,
      palletShare: inventoryUnits ? palletUnits / inventoryUnits : 0,
      shelfLocations,
      shelfBins,
      binsPerLocation,
      onePerBin,
      // Room that stays empty because articles do not share bins.
      binsWasted: onePerBin ? Math.max(0, shelfBins * unitsPerBin - shelfUnits) : 0,
      drawerLocations,
      pallets,
      capacity,
      capacityShortfall: Math.max(0, inventoryUnits - capacity),
    };
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

  // ---- Consolidated table sorting -------------------------------------------------------------------
  // Sorts a list of consolidated items by one column's sort key. Numeric columns compare as numbers so
  // "1 200" orders above "80"; text columns use a locale-aware numeric collation so "A2" precedes "A10".
  // Ties fall back to `tieBreaker` so the order is stable between renders.
  function compareSortValues(left, right, collatorOptions = { numeric: true, sensitivity: 'base' }) {
    const leftBlank = left === null || left === undefined || left === '';
    const rightBlank = right === null || right === undefined || right === '';
    // Blanks always sink to the bottom, whichever direction the column is sorted in.
    if (leftBlank && rightBlank) return 0;
    if (leftBlank) return 1;
    if (rightBlank) return -1;
    const leftNumber = Number(left);
    const rightNumber = Number(right);
    if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) return leftNumber - rightNumber;
    const collator = new Intl.Collator(undefined, collatorOptions);
    return collator.compare(String(left), String(right));
  }

  // `sort` is null (default order), or { key, direction } with direction 'asc' or 'desc'.
  // An unknown key returns the list untouched, so a stale saved sort can never blank the table.
  function sortItems(items, sort, options = {}) {
    const list = Array.isArray(items) ? items : [];
    const key = sort?.key;
    if (!key || !['asc', 'desc'].includes(sort?.direction)) return list;
    const keyFor = options.keyFor || ((item) => item?.[key]);
    const tieBreaker = options.tieBreaker || ((item) => item?.key);
    const sign = sort.direction === 'desc' ? -1 : 1;
    const isBlank = (value) => value === null || value === undefined || value === '';
    return [...list].sort((a, b) => {
      const left = keyFor(a);
      const right = keyFor(b);
      // Rows with nothing to compare against stay at the bottom whichever way the column is sorted,
      // so the descending direction does not push incomplete rows to the top.
      if (isBlank(left) || isBlank(right)) {
        if (isBlank(left) && isBlank(right)) return 0;
        return isBlank(left) ? 1 : -1;
      }
      const primary = compareSortValues(left, right);
      if (primary !== 0) return primary * sign;
      const tie = compareSortValues(tieBreaker(a), tieBreaker(b));
      return tie * sign;
    });
  }

  // ---- Missing-input classification ------------------------------------------------------------------
  // Which of a row's values are still blank, and why. Kept separate from the table so the summary card,
  // the filter and the "what to do about it" hint all read from one definition.
  const MISSING_PRICE = 'price';
  const MISSING_FREIGHT = 'freight';

  function missingInputReasons(item) {
    const reasons = [];
    if (item?.unitPrice === null || item?.unitPrice === undefined) reasons.push(MISSING_PRICE);
    if (item?.fx === null || item?.fx === undefined) reasons.push(MISSING_PRICE);
    if (item?.missingFreight === true) reasons.push(MISSING_FREIGHT);
    return [...new Set(reasons)];
  }

  /* ---------- Sales history: how many pieces of each article a market has bought ---------- */

  // Article numbers are compared without case or spaces, so "2513-ag-11 " matches "2513-AG-11".
  function historyKey(value) {
    return cellText(value).toUpperCase().replace(/\s+/g, '');
  }

  // Quantities arrive as "321.00", "1 234,5" or plain numbers.
  function historyNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    let text = cellText(value).replace(/[\s ]/g, '');
    if (!/^-?[\d.,]+$/.test(text)) return null;
    const comma = text.lastIndexOf(',');
    const dot = text.lastIndexOf('.');
    if (comma >= 0 && dot >= 0) text = comma > dot ? text.replace(/\./g, '').replace(',', '.') : text.replace(/,/g, '');
    else if (comma >= 0) text = text.length - comma - 1 <= 2 ? text.replace(',', '.') : text.replace(/,/g, '');
    const number = Number(text);
    return Number.isFinite(number) ? number : null;
  }

  // Reads a purchase/sales history: one column with the article number (optionally followed by " - description"),
  // an optional description column and a quantity column. Several lines for the same article are added up.
  // Returns null when the rows do not look like such a history (price lists are not histories).
  function parseSalesHistory(rows) {
    if (!Array.isArray(rows)) return null;
    let headerRow = -1;
    let itemColumn = -1;
    let quantityColumn = -1;
    let descriptionColumn = -1;
    for (let index = 0; index < Math.min(rows.length, 12) && headerRow < 0; index += 1) {
      const cells = (rows[index] || []).map((cell) => cellText(cell).toLowerCase());
      if (cells.some((cell) => /price|pris|net\b|discount|rabatt/.test(cell))) continue;
      const item = cells.findIndex((cell) => /^(?:item|article|part|artikel|art\.?)(?:\b|\s|$)/.test(cell));
      const quantity = cells.findIndex((cell, column) => column !== item && /^(?:quantity|qty|antal|kvantitet|mengde|sold|units)\b/.test(cell));
      if (item < 0 || quantity < 0) continue;
      headerRow = index;
      itemColumn = item;
      quantityColumn = quantity;
      descriptionColumn = cells.findIndex((cell, column) => column !== item && /^(?:description|beskrivning|name|artikelnamn|benevnelse)\b/.test(cell));
    }
    if (headerRow < 0) return null;
    // A second header line with price columns means this is a price list, not a history.
    if ((rows[headerRow + 1] || []).some((cell) => /^(?:list |net |unit |total )?(?:price|pris)\b/i.test(cellText(cell)))) return null;
    const combined = descriptionColumn < 0;
    const merged = new Map();
    let reportedTotal = null;
    let lines = 0;
    for (let index = headerRow + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      const itemText = cellText(row[itemColumn]);
      const quantity = historyNumber(row[quantityColumn]);
      if (!itemText || quantity === null) continue;
      if (/^(?:total|totalt|summa|sum)\b/i.test(itemText)) {
        reportedTotal = quantity;
        continue;
      }
      let itemNo = itemText;
      let description = combined ? '' : cellText(row[descriptionColumn]);
      if (combined) {
        const split = itemText.match(/^(.+?)\s+-\s+(.*)$/);
        if (split) {
          itemNo = split[1].trim();
          description = split[2].trim();
        }
      }
      const key = historyKey(itemNo);
      if (!key) continue;
      lines += 1;
      const entry = merged.get(key) || { key, itemNo, description, quantity: 0, descriptionQuantity: -Infinity };
      entry.quantity += quantity;
      // The description of the line with the most pieces is the one shown.
      if (quantity > entry.descriptionQuantity && description) {
        entry.description = description;
        entry.descriptionQuantity = quantity;
      }
      merged.set(key, entry);
    }
    const items = [...merged.values()]
      .filter((entry) => entry.quantity > 0)
      .map(({ key, itemNo, description, quantity }) => ({ key, itemNo, description, quantity }))
      .sort((a, b) => b.quantity - a.quantity || a.itemNo.localeCompare(b.itemNo, undefined, { numeric: true }));
    if (!items.length) return null;
    return { items, lines, total: items.reduce((sum, entry) => sum + entry.quantity, 0), reportedTotal };
  }

  // Competition ranking: equal values share a rank (1, 2, 2, 4). Values of zero or less are not ranked.
  function rankValues(values) {
    const sorted = [...values].filter((value) => value > 0).sort((a, b) => b - a);
    return values.map((value) => (value > 0 ? sorted.indexOf(value) + 1 : null));
  }

  // Joins the parts of the project to the sales of each market and ranks them.
  //   parts:   [{ key, part, altParts, description, kits: [kit names], manufacturer }]
  //   markets: [{ name, items: [{ key, itemNo, description, quantity }] }]
  // A part is found by its article number or by its alternative number; one sale is never counted twice.
  function buildPopularity({ parts = [], markets = [], kitOrder = [] } = {}) {
    const lookups = markets.map((market) => new Map((market.items || []).map((entry) => [entry.key, entry])));
    const matchedKeys = markets.map(() => new Set());
    const rows = parts.map((part) => {
      const keys = [...new Set([part.part, ...(part.altParts || [])].map(historyKey).filter(Boolean))];
      const sold = lookups.map((lookup, marketIndex) => keys.reduce((sum, key) => {
        const entry = lookup.get(key);
        if (!entry) return sum;
        matchedKeys[marketIndex].add(key);
        return sum + entry.quantity;
      }, 0));
      return { key: part.key, part: part.part, altParts: part.altParts || [], description: part.description || '', kits: part.kits || [], manufacturer: part.manufacturer || '', inKit: true, sold };
    });
    // Articles that were bought but are in no kit are kept: they show what the kits do not cover.
    const kitKeys = new Set(rows.flatMap((row) => [row.part, ...row.altParts].map(historyKey)));
    const outside = new Map();
    markets.forEach((market, marketIndex) => {
      (market.items || []).forEach((entry) => {
        if (kitKeys.has(entry.key)) return;
        const row = outside.get(entry.key) || { key: entry.key, part: entry.itemNo, altParts: [], description: entry.description, kits: [], manufacturer: '', inKit: false, sold: markets.map(() => 0) };
        row.sold[marketIndex] += entry.quantity;
        if (!row.description) row.description = entry.description;
        outside.set(entry.key, row);
      });
    });
    outside.forEach((row) => rows.push(row));

    const totals = rows.map((row) => row.sold.reduce((sum, value) => sum + value, 0));
    const totalRanks = rankValues(totals);
    const marketRanks = markets.map((market, index) => rankValues(rows.map((row) => row.sold[index])));
    rows.forEach((row, index) => {
      row.total = totals[index];
      row.rank = totalRanks[index];
      row.marketRanks = markets.map((market, marketIndex) => marketRanks[marketIndex][index]);
      row.inAll = markets.length > 0 && row.sold.every((value) => value > 0);
      row.weakest = markets.length ? Math.min(...row.sold) : 0;
      row.bothRank = null;
    });
    // "Popular in both": ranked by the market that buys the least, so one big market cannot carry a part alone.
    rows.filter((row) => row.inAll)
      .sort((a, b) => b.weakest - a.weakest || b.total - a.total)
      .forEach((row, index, list) => {
        const previous = list[index - 1];
        row.bothRank = previous && previous.weakest === row.weakest && previous.total === row.total ? previous.bothRank : index + 1;
      });
    rows.sort((a, b) => b.total - a.total || a.part.localeCompare(b.part, undefined, { numeric: true }));

    const kitNames = kitOrder.length ? kitOrder : [...new Set(rows.flatMap((row) => row.kits))];
    const kits = kitNames.map((name) => {
      const members = rows.filter((row) => row.inKit && row.kits.includes(name));
      const sold = markets.map((market, index) => members.reduce((sum, row) => sum + row.sold[index], 0));
      const top = members.find((row) => row.total > 0) || null;
      return {
        name,
        parts: members.length,
        soldParts: members.filter((row) => row.total > 0).length,
        sold,
        total: sold.reduce((sum, value) => sum + value, 0),
        top: top ? { part: top.part, description: top.description, total: top.total } : null,
        members,
      };
    }).sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, undefined, { numeric: true }));
    const kitRanks = rankValues(kits.map((kit) => kit.total));
    kits.forEach((kit, index) => { kit.rank = kitRanks[index]; });

    const coverage = markets.map((market, index) => {
      const entries = market.items || [];
      const matched = entries.filter((entry) => matchedKeys[index].has(entry.key));
      return {
        name: market.name,
        units: entries.reduce((sum, entry) => sum + entry.quantity, 0),
        matchedUnits: matched.reduce((sum, entry) => sum + entry.quantity, 0),
        articles: entries.length,
        matchedArticles: matched.length,
      };
    });
    return { rows, kits, coverage };
  }

  /* ---------- Suggested stock from the sales history ---------- */

  // z-values for the share of the demand that should be covered without running out.
  const SERVICE_LEVELS = { 0.9: 1.282, 0.95: 1.645, 0.99: 2.326 };

  // Suggests how many pieces of each kit part to keep in stock.
  //   rows:   the kit rows of buildPopularity (inKit, total, key, …); kits: its kit list (members, parts, soldParts)
  //   a part that sells needs: its demand over the cover period, plus a safety margin for the chance that
  //   demand runs above average (demand is treated as random "arrivals", so the margin grows with the square
  //   root of the expected demand), and at least one piece.
  //   unsold parts get nothing, except in kits that are mostly sold (completeKits) where one piece keeps the kit complete.
  //   budget (optional, in the same money as unitCost): trims the plan in steps until it fits:
  //   safety margins first, then the slow sellers down to one piece, then the kit-completing pieces, then the slowest parts.
  function suggestStock({ rows = [], kits = [], years = 5, monthsCover = 6, serviceLevel = 0.95, completeKits = true, budget = null, binSize = 0, unitCost = () => null, current = () => null } = {}) {
    const z = SERVICE_LEVELS[serviceLevel] ?? SERVICE_LEVELS[0.95];
    const spanYears = Math.max(years, 0.1);
    const parts = rows.filter((row) => row.inKit);
    const sold = parts.filter((row) => row.total > 0).sort((a, b) => b.total - a.total);
    const grand = sold.reduce((sum, row) => sum + row.total, 0);
    const classes = new Map();
    let running = 0;
    sold.forEach((row) => {
      classes.set(row.key, grand <= 0 || running / grand < 0.8 ? 'A' : running / grand < 0.95 ? 'B' : 'C');
      running += row.total;
    });
    const completers = new Set();
    if (completeKits) {
      kits.forEach((kit) => {
        if (kit.parts > 0 && kit.soldParts / kit.parts >= 0.5) kit.members.filter((member) => member.total === 0).forEach((member) => completers.add(member.key));
      });
    }
    const lines = parts.map((row) => {
      const perYear = row.total / spanYears;
      const cover = (perYear * monthsCover) / 12;
      const safety = row.total > 0 ? z * Math.sqrt(cover) : 0;
      const wanted = row.total > 0 ? Math.max(1, Math.ceil(cover + safety)) : completers.has(row.key) ? 1 : 0;
      // With one article per bin, a part-filled bin is paid for anyway: the stock is filled up to whole bins.
      const needed = binSize > 1 && wanted > 0 ? Math.ceil(wanted / binSize) * binSize : wanted;
      const cost = unitCost(row);
      return {
        row,
        key: row.key,
        cls: classes.get(row.key) || '',
        perYear,
        cover,
        safety,
        target: needed,
        reason: row.total > 0 ? 'sales' : needed ? 'kit' : 'none',
        cost: Number.isFinite(cost) ? cost : null,
        now: current(row),
      };
    });
    const valueOf = () => lines.reduce((sum, line) => sum + (line.cost === null ? 0 : line.cost * line.target), 0);
    const limit = Number.isFinite(budget) && budget > 0 ? budget : null;
    if (limit !== null && valueOf() > limit) {
      const steps = [
        () => lines.forEach((line) => {
          const base = line.row.total > 0 ? Math.max(1, Math.ceil(line.cover)) : line.target;
          if (base < line.target) { line.target = base; line.reason = 'trimmed'; }
        }),
        () => lines.forEach((line) => {
          if (line.cls === 'C' && line.target > 1) { line.target = 1; line.reason = 'trimmed'; }
        }),
        () => lines.forEach((line) => {
          if (line.row.total === 0 && line.target > 0) { line.target = 0; line.reason = 'cut'; }
        }),
      ];
      // The slowest sellers (class C) are dropped before the popular parts are reduced.
      steps.push(() => lines.forEach((line) => {
        if (line.cls === 'C' && line.target > 0) { line.target = 0; line.reason = 'cut'; }
      }));
      // The popular parts then share the cut in proportion, keeping at least one piece each.
      steps.push(() => {
        const factor = limit / valueOf();
        lines.forEach((line) => {
          if (line.target > 1) { line.target = Math.max(1, Math.floor(line.target * factor)); line.reason = 'trimmed'; }
        });
      });
      for (const step of steps) {
        if (valueOf() <= limit) break;
        step();
      }
      if (valueOf() > limit) {
        [...lines].filter((line) => line.target > 0 && line.cost !== null).sort((a, b) => a.row.total - b.row.total).forEach((line) => {
          if (valueOf() <= limit) return;
          line.target = 0;
          line.reason = 'cut';
        });
      }
    }
    lines.forEach((line) => {
      line.value = line.cost === null ? null : line.cost * line.target;
      line.valueNow = line.cost === null || line.now === null ? null : line.cost * line.now;
    });
    lines.sort((a, b) => b.row.total - a.row.total || a.row.part.localeCompare(b.row.part, undefined, { numeric: true }));
    return {
      lines,
      value: valueOf(),
      valueNow: lines.reduce((sum, line) => sum + (line.valueNow ?? 0), 0),
      pieces: lines.reduce((sum, line) => sum + line.target, 0),
      stocked: lines.filter((line) => line.target > 0).length,
      overBudget: limit !== null && valueOf() > limit,
      unknownCost: lines.filter((line) => line.target > 0 && line.cost === null).length,
    };
  }

  function hasMissingInputs(item) {
    return missingInputReasons(item).length > 0;
  }

  return {
    assessProfitabilityCompleteness,
    buildPopularity,
    historyKey,
    parseSalesHistory,
    allocationShares,
    analyzeBins,
    binsForUnits,
    compareSortValues,
    EU_MEMBERS,
    catalogueSheetName,
    catalogueSheetRows,
    clampWirePosition,
    parseKitCatalogue,
    curveAroundRect,
    fitProjection,
    hasMissingInputs,
    placeBox,
    quadraticBow,
    rayToRect,
    rectAround,
    rectContains,
    rectsOverlap,
    customsTreatment,
    importBounds,
    missingInputReasons,
    normalizeCountryCode,
    planWarehouseStorage,
    projectScenario,
    requiredLocations,
    salesVatRate,
    sortItems,
    suggestStock,
    sanitizeWirePositions,
    summarizeOccurrences,
  };
}));
