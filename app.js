/* global XLSX, HyperFormula, PartsListCore */

(() => {
  'use strict';

  const FormulaEngine = window.HyperFormula?.HyperFormula || window.HyperFormula;
  const Core = window.PartsListCore;
  const MAX_IMPORT_ROWS = 5000;
  const MAX_IMPORT_COLS = 100;
  const MAX_GRID_ROWS = 250;
  const MAX_GRID_COLS = 30;
  const MAX_FILL_CELLS = 10000;
  const RATE_CACHE_MS = 6 * 60 * 60 * 1000;
  const RATE_SOURCE = 'https://api.frankfurter.dev/v2';
  const WAREHOUSE_RATE_STORAGE_KEY = 'partslist.warehouse-rates.caesar14.v1';
  const WAREHOUSE_RATE_PRIVATE_FILE = 'warehouse-rates.private.json';
  // Parts lists saved encrypted in the repository (see vault.js and scripts/encrypt-documents.js).
  const VAULT_FILE = 'data/bundle.dat';
  const UI_PREFERENCES_KEY = 'partslist.ui-preferences';
  // Undo keeps the last few reversible edits. Snapshots are small (settings + item overrides + variables),
  // not whole workbooks, so a bounded stack stays cheap.
  const UNDO_LIMIT = 40;
  // Typing in a number field must not rebuild the whole table on every keystroke.
  const RECALC_DEBOUNCE_MS = 180;
  // Columns that are rarely needed in day-to-day use. Hidden by default so the table fits a normal
  // screen, and reachable in one click from the Columns menu.
  const DEFAULT_HIDDEN_COLUMNS = Object.freeze(['fx', 'convertedUnit', 'discount', 'sellingTotal', 'importCosts', 'shippingMargin', 'pricingSource']);
  const CAESAR_SHIFT = 14;
  const PRINTABLE_ASCII_START = 32;
  const PRINTABLE_ASCII_RANGE = 95;
  const CURRENCIES = ['SEK', 'EUR', 'USD', 'GBP', 'NOK', 'DKK', 'CHF', 'CAD', 'AUD', 'JPY', 'CNY', 'PLN'];
  const WAREHOUSE_DEFAULTS = Object.freeze({
    country: '578',
    shelfEnabled: true,
    drawerEnabled: false,
    palletEnabled: false,
    plannedPallets: 0,
    shelfShare: 55,
    drawerShare: 25,
    binsPerShelf: 4,
    shelvesPerRack: 4,
    unitsPerBin: 5,
    unitsPerShelf: 20,
    unitsPerDrawer: 8,
    unitsPerPallet: 60,
    palletType: 'eu',
    palletHeight: '120',
    receipts: 2,
    receiptLines: 8,
    orders: 10,
    orderLines: 5,
    ediLabels: 10,
    businessParcels: 8,
    privateParcels: 2,
    packaging: 0,
    includeWms: false,
    site: 'Norwegian 3PL',
  });
  const FREIGHT_DEFAULTS = Object.freeze({
    consolidatedShipment: 0,
    consolidatedCurrency: 'USD',
    allocation: 'value',
    dutyRate: 0,
    insuranceRate: 0,
    clearanceFee: 0,
    clearanceCurrency: 'NOK',
  });
  const SALES_DEFAULTS = Object.freeze({
    defaultMode: 'turns',
    defaultValue: 1,
    horizonMonths: 24,
    customerPaysOutbound: false,
  });
  const KIT_DEFAULTS = Object.freeze({
    shippingMode: 'sheet', shippingAmount: 1, shippingCurrency: 'SEK', manufacturer: '', originCountry: '840', defaultCurrency: null, discountRate: null, dutyRate: null, fxMargin: null,
  });
  const KIT_SHIPPING_MODES = { sheet: 'Sheet values', perLine: 'Per line', kitTotal: 'Whole kit', percent: 'Percent of item value' };
  // Settings that belong to a manufacturer's route as a whole rather than to one kit sheet. A manufacturer either
  // shares the project-wide consolidated shipment ('project') or ships on its own ('own').
  const MAKER_DEFAULTS = Object.freeze({ location: '', shipmentMode: 'project', shipmentAmount: 0, shipmentCurrency: 'SEK', clearanceAmount: 0, clearanceCurrency: 'SEK', insuranceRate: null });
  const SUPPLIER_DEFAULTS = Object.freeze({ name: 'TEI Rock Drills', location: 'Montrose, Colorado', country: '840', countryName: 'United States', lat: 38.478, lon: -107.876 });
  // ISO 3166-1 numeric codes, matching the world-atlas country ids.
  const EU_MEMBERS = new Set(Core.EU_MEMBERS);
  const EFTA_MEMBERS = new Set(['352', '438', '756']);
  const ORIGIN_PRESETS = Object.freeze({
    578: { name: 'Norway', place: 'Oslo', lat: 59.913, lon: 10.752 },
    840: { name: 'United States', place: 'Montrose, Colorado', lat: 38.478, lon: -107.876 },
    756: { name: 'Switzerland', place: 'Zürich', lat: 47.377, lon: 8.54 },
    100: { name: 'Bulgaria', place: 'Sofia', lat: 42.698, lon: 23.322 },
    276: { name: 'Germany', place: 'Frankfurt', lat: 50.11, lon: 8.682 },
    752: { name: 'Sweden', place: 'Stockholm', lat: 59.329, lon: 18.069 },
    380: { name: 'Italy', place: 'Milan', lat: 45.464, lon: 9.19 },
    826: { name: 'United Kingdom', place: 'London', lat: 51.507, lon: -0.128 },
    124: { name: 'Canada', place: 'Toronto', lat: 43.653, lon: -79.383 },
    156: { name: 'China', place: 'Shanghai', lat: 31.23, lon: 121.474 },
    392: { name: 'Japan', place: 'Tokyo', lat: 35.676, lon: 139.65 },
  });
  const DESTINATIONS = Object.freeze({
    578: {
      name: 'Norway', localName: 'Norge', vat: 0.25, arrival: [11.1, 60.2], warehouse: [10.95, 59.7],
      bounds: [2.5, 57.6, 32, 71.4],
      customers: [['Bergen', 60.391, 5.322], ['Ålesund', 62.472, 6.155], ['Stavanger', 58.97, 5.733], ['Kristiansand', 58.146, 7.996], ['Trondheim', 63.431, 10.395], ['Bodø', 67.28, 14.405], ['Tromsø', 69.649, 18.955]],
    },
    752: {
      name: 'Sweden', localName: 'Sverige', vat: 0.25, arrival: [11.973, 57.708], warehouse: [18.069, 59.329],
      bounds: [10.5, 54.3, 24.8, 69.4],
      customers: [['Malmö', 55.605, 13.003], ['Göteborg', 57.708, 11.974], ['Örebro', 59.275, 15.214], ['Stockholm', 59.329, 18.069], ['Sundsvall', 62.391, 17.306], ['Umeå', 63.826, 20.263], ['Luleå', 65.584, 22.154]],
    },
  });
  // Currencies that name one country unambiguously; used only to pre-fill the origin of an added kit.
  const CURRENCY_HOME_COUNTRY = Object.freeze({ USD: '840', CHF: '756', GBP: '826', JPY: '392', CNY: '156', CAD: '124', NOK: '578', SEK: '752' });
  const MAP_DATA_URL = 'https://cdn.jsdelivr.net/npm/world-atlas@2.0.2/countries-50m.json';
  const MAP_STUB = 34;                 // wire kept visible on each side of a table so the arrow reads as passing through it
  const MAP_OVERLAY_MIN_WIDTH = 820;   // narrower than this the tables sit under the map instead of on it
  // How VAT is treated for sales into each country: 'none', or 'price' = added to the price after margin.
  const VAT_DEFAULTS = Object.freeze({ 578: 'none', 752: 'price' });
  const VAT_TREATMENT_KEYS = ['none', 'price'];
  const VAT_LABELS = { none: 'No VAT', price: 'VAT on the price after margin' };
  const FLOW_OVERRIDE_KEYS = ['outboundPerUnit', 'outtakePerUnit', 'storagePerUnitMonth'];
  // Validated categorical order (dataviz reference palette); groups take slots in creation order.
  const GROUP_COLORS = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948'];
  const MAX_GROUPS = GROUP_COLORS.length;
  const LEGACY_SETTINGS_STORAGE_PREFIX = 'partslist.settings.caesar14.v1:';
  const SETTINGS_SHEET_MARKER = 'partslist settings';
  const VIEW = Object.freeze({ consolidated: ':consolidated', dashboard: ':dashboard', warehouse: ':warehouse', wire: ':wire' });
  const SPECIAL_VIEWS = new Set(Object.values(VIEW));
  const WAREHOUSE_RATE_DEFAULTS = Object.freeze({
    eu120: 0,
    eu220: 0,
    sea120: 0,
    sea220: 0,
    shelf: 0,
    drawer: 0,
    edi: 0,
    receiptBase: 0,
    receiptLine: 0,
    orderBase: 0,
    orderLine: 0,
    parcel: 0,
    privateSurcharge: 0,
    wms: 0,
  });

  const FIELD_LABELS = {
    part: 'Part / SKU',
    altPart: 'Alt. part no.',
    description: 'Description',
    quantity: 'Quantity',
    price: 'Unit price',
    discountedTotal: 'Discounted total',
    sellingTotal: 'Price after multiplier',
    shipping: 'Shipping',
    shippingWithMargin: 'Shipping incl. margin',
    currency: 'Currency',
  };

  // ---------- Language ----------
  // The interface is written in English; the Swedish pack (i18n-sv.js) translates what is rendered.
  // Each original string is remembered per text node and attribute, so switching back is lossless and
  // anything the app rewrites later (status messages, tooltips, table cells) is translated again.
  const I18n = window.PartsListI18n;
  const TRANSLATED_ATTRIBUTES = ['title', 'placeholder', 'aria-label', 'data-tip', 'alt'];
  const ATTRIBUTE_SELECTOR = TRANSLATED_ATTRIBUTES.map((name) => `[${name}]`).join(',');
  const SKIP_TRANSLATION = 'script, style, textarea, [translate="no"]';
  const textRecords = new WeakMap();
  const attributeRecords = new WeakMap();

  function translateUi(text) {
    return state.language === 'sv' && I18n ? I18n.translate('sv', text) : text;
  }

  function localizeTextNode(node) {
    if (!node.nodeValue || node.parentElement?.closest(SKIP_TRANSLATION)) return;
    let record = textRecords.get(node);
    // Text the app wrote itself becomes the new source; text we wrote matches record.shown.
    if (!record || node.nodeValue !== record.shown) record = { source: node.nodeValue, shown: node.nodeValue };
    const shown = translateUi(record.source);
    if (shown !== node.nodeValue) node.nodeValue = shown;
    record.shown = shown;
    textRecords.set(node, record);
  }

  function localizeAttribute(element, attribute) {
    if (!element.hasAttribute(attribute) || element.closest(SKIP_TRANSLATION)) return;
    let records = attributeRecords.get(element);
    if (!records) { records = {}; attributeRecords.set(element, records); }
    const current = element.getAttribute(attribute);
    let record = records[attribute];
    if (!record || current !== record.shown) record = { source: current, shown: current };
    const shown = translateUi(record.source);
    if (shown !== current) element.setAttribute(attribute, shown);
    record.shown = shown;
    records[attribute] = record;
  }

  function localizeSubtree(root) {
    if (!root) return;
    if (root.nodeType === Node.TEXT_NODE) {
      localizeTextNode(root);
      return;
    }
    if (root.nodeType !== Node.ELEMENT_NODE || root.closest(SKIP_TRANSLATION)) return;
    [root, ...root.querySelectorAll(ATTRIBUTE_SELECTOR)].forEach((element) => {
      TRANSLATED_ATTRIBUTES.forEach((attribute) => localizeAttribute(element, attribute));
    });
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) localizeTextNode(walker.currentNode);
  }

  function observeLanguage() {
    const observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        if (mutation.type === 'childList') mutation.addedNodes.forEach(localizeSubtree);
        else if (mutation.type === 'characterData') localizeTextNode(mutation.target);
        else if (mutation.type === 'attributes') localizeAttribute(mutation.target, mutation.attributeName);
      });
    });
    observer.observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: TRANSLATED_ATTRIBUTES });
  }

  function setLanguage(language, rerender = true) {
    state.language = language === 'en' ? 'en' : 'sv';
    document.documentElement.lang = state.language;
    document.title = translateUi('PartsList Workbook');
    dom.languageSelect.querySelectorAll('[data-language]').forEach((button) => button.setAttribute('aria-pressed', String(button.dataset.language === state.language)));
    localizeSubtree(document.body);
    populateOriginSelects();
    populateCustomerSelects();
    if (rerender && state.workbook) {
      renderMappings();
      renderTabs();
      refreshViews();
    }
  }

  const HEADER_TERMS = {
    part: ['part number', 'part no', 'part #', 'part', 'sku', 'item code', 'item number', 'item no', 'article number', 'article', 'component', 'product code', 'product', 'artikelnummer', 'artikelnr', 'reservdelsnummer', 'teilenummer', 'artikel nr', 'numero de piece', 'référence'],
    altPart: ['alt article no', 'alt part no', 'alt article number', 'alt part number', 'alternative article number', 'alternative part number', 'alternate article number', 'alternate part number', 'secondary article number', 'secondary part number'],
    description: ['part description', 'item description', 'description', 'details', 'name', 'beskrivning', 'beteckning', 'bezeichnung', 'benämning', 'наименование'],
    quantity: ['maximum quantity', 'max quantity', 'quantity', 'qty', 'amount', 'count', 'pcs', 'pieces', 'units', 'antal', 'mängd', 'menge', 'quantite', 'quantité', 'количество'],
    price: ['unit price', 'unit cost', 'purchase price', 'price each', 'price', 'cost each', 'cost', 'styckpris', 'inköpspris', 'pris', 'einzelpreis', 'preis', 'prix unitaire', 'цена'],
    discountedTotal: ['rabatt 30', 'rabatt', 'discounted total', 'after discount', 'discount'],
    sellingTotal: ['dubblering', 'price after multiplier', 'selling total', 'sales total'],
    shipping: ['frakt', 'shipping cost', 'shipment cost', 'freight cost', 'delivery cost', 'transport cost', 'shipping', 'shipment', 'freight', 'delivery', 'transport'],
    shippingWithMargin: ['inkl fraktmarginal', 'fraktmarginal', 'shipping incl margin', 'freight incl margin', 'shipping with margin'],
    currency: ['currency code', 'currency', 'curr', 'ccy', 'valuta', 'währung', 'devise', 'валута'],
  };

  const state = {
    language: 'sv',
    workbook: null,
    fileName: '',
    matrices: {},
    hf: null,
    sheetIds: {},
    mappings: {},
    variables: [],
    consolidated: [],
    rates: {},
    rateDates: {},
    pairRates: {},
    liveRateCells: [],
    dependencyIndex: new Map(),
    activeView: VIEW.consolidated,
    selectedCell: null,
    config: {
      discount: 0.30,
      multiplier: 2,
      shippingMargin: 0.15,
      vatRate: 0,
      sourceCurrency: 'USD',
      outputCurrency: 'SEK',
    },
    warehouse: {
      ...WAREHOUSE_DEFAULTS,
      rates: { ...WAREHOUSE_RATE_DEFAULTS },
    },
    freight: { ...FREIGHT_DEFAULTS },
    customer: { country: '578' },
    vat: { unlocked: false, byCountry: { ...VAT_DEFAULTS } },
    kits: {},
    makers: {},
    laneRefs: new Map(),
    groups: [],
    items: {},
    sales: { ...SALES_DEFAULTS },
    scenario: { salesDelta: 0, purchaseDelta: 0, inboundDelta: 0, warehouseDelta: 0 },
    supplier: { ...SUPPLIER_DEFAULTS },
    flowOverrides: { outboundPerUnit: null, outtakePerUnit: null, storagePerUnitMonth: null },
    flowUnlocked: false,
    flowMode: 'route',
    map: { features: null, names: null, loading: false, error: null, compare: null, hover: null, pinned: null, open: { warehouse: false, customer: false }, worldKey: null, nordicKey: null, landKey: null, land: null },
    customsPinned: false,
    excludedItems: [],
    kitSummaries: [],
    selectedKeys: new Set(),
    selectionAnchor: null,
    selectionAction: 'group',
    sheetSelection: { sheetName: null, rows: new Set(), anchor: null },
    consolidatedSearch: '',
    consolidatedFilter: 'all',
    consolidatedSort: null,
    hiddenColumns: [...DEFAULT_HIDDEN_COLUMNS],
    focusedRowKey: null,
    undoStack: [],
    redoStack: [],
    undoLabel: '',
    vaultBytes: null,
    recalcTimer: null,
    pendingRecalc: null,
    suppressHistory: false,
    suppressHash: false,
    pendingDisabledKits: null,
    pendingMappings: null,
    pendingVariables: null,
    pendingWireRelations: null,
    importWarnings: [],
    wireRelations: [],
    wireNodePositions: { overview: {}, underhood: {} },
    wireDrag: null,
    suppressWireClickUntil: 0,
    selectedWireNode: 'consolidated.quantity',
    wireMode: 'overview',
    wireExpanded: false,
    guideDismissed: false,
    projectDirty: false,
    projectSaveLabel: '',
    sidebarCollapsed: false,
    warehousePreviewTransform: { scale: 1, x: 0, y: 0 },
    warehousePreviewDrag: null,
    secureDownloadUrl: null,
  };

  const el = (id) => document.getElementById(id);
  const dom = {
    file: el('workbook-file'),
    addKitFile: el('add-kit-file'),
    vaultForm: el('vault-form'),
    vaultKey: el('vault-key'),
    vaultOpen: el('vault-open'),
    addKitLabel: el('add-kit-label'),
    languageSelect: el('language-select'),
    sidebarBody: el('sidebar-body'),
    sidebarCollapse: el('sidebar-collapse'),
    sidebarExpand: el('sidebar-expand'),
    exportButton: el('export-button'),
    undoButton: el('undo-button'),
    columnMenu: el('column-menu'),
    columnList: el('column-list'),
    columnsAll: el('columns-all'),
    columnsReset: el('columns-reset'),
    columnsDone: el('columns-done'),
    saveState: el('save-state'),
    workspace: el('workspace'),
    sidebar: el('app-sidebar'),
    sidebarToggle: el('sidebar-toggle'),
    guideToggle: el('guide-toggle'),
    statusBar: el('status-bar'),
    statusMessage: el('status-message'),
    emptyPanel: el('empty-panel'),
    controls: el('workbook-controls'),
    workbookName: el('workbook-name'),
    sheetCount: el('sheet-count'),
    sourceCurrency: el('source-currency'),
    outputCurrency: el('output-currency'),
    warehouseCountry: el('warehouse-country'),
    refreshRates: el('refresh-rates'),
    rateStatus: el('rate-status'),
    mappingList: el('mapping-list'),
    variablesList: el('variables-list'),
    variableForm: el('variable-form'),
    variableName: el('variable-name'),
    variableExpression: el('variable-expression'),
    securePassword: el('secure-password'),
    securePasswordConfirm: el('secure-password-confirm'),
    secureExport: el('secure-export'),
    secureFile: el('secure-project-file'),
    secureDownload: el('secure-download'),
    tabs: el('view-tabs'),
    setupGuide: el('setup-guide'),
    setupGuideSummary: el('setup-guide-summary'),
    setupProgress: document.querySelector('.setup-progress'),
    setupProgressBar: el('setup-progress-bar'),
    setupChecklist: el('setup-checklist'),
    guideClose: el('guide-close'),
    welcome: el('welcome-view'),
    consolidatedView: el('consolidated-view'),
    consolidatedSummary: el('consolidated-summary'),
    summaryCards: el('summary-cards'),
    consolidatedTable: el('consolidated-table'),
    warehouseView: el('warehouse-view'),
    warehouseDataNote: el('warehouse-data-note'),
    warehouseSummaryCards: el('warehouse-summary-cards'),
    warehousePalletShare: el('warehouse-pallet-share'),
    warehouseShareWarning: el('warehouse-share-warning'),
    warehouseCapacityLabel: el('warehouse-capacity-label'),
    warehousePreview: el('warehouse-preview'),
    warehouseStorageLegend: el('warehouse-storage-legend'),
    warehouseCostBody: el('warehouse-cost-body'),
    warehouseMonthlyTotal: el('warehouse-monthly-total'),
    warehouseOutputHeader: el('warehouse-output-header'),
    warehouseCurrencyReference: el('warehouse-currency-reference'),
    warehouseRateInputs: el('warehouse-rate-inputs'),
    warehouseRatesExport: el('warehouse-rates-export'),
    warehouseRatesFile: el('warehouse-rates-file'),
    warehouseRatesStatus: el('warehouse-rates-status'),
    whShelfShare: el('wh-shelf-share'),
    whDrawerShare: el('wh-drawer-share'),
    whShelfEnabled: el('wh-shelf-enabled'),
    whDrawerEnabled: el('wh-drawer-enabled'),
    whPalletEnabled: el('wh-pallet-enabled'),
    whPlannedPallets: el('wh-planned-pallets'),
    whBinsPerShelf: el('wh-bins-per-shelf'),
    whShelvesPerRack: el('wh-shelves-per-rack'),
    whUnitsPerBin: el('wh-units-per-bin'),
    whUnitsShelf: el('wh-units-shelf'),
    whUnitsDrawer: el('wh-units-drawer'),
    whUnitsPallet: el('wh-units-pallet'),
    whPalletType: el('wh-pallet-type'),
    whPalletHeight: el('wh-pallet-height'),
    whReceipts: el('wh-receipts'),
    whReceiptLines: el('wh-receipt-lines'),
    whOrders: el('wh-orders'),
    whOrderLines: el('wh-order-lines'),
    whEdiLabels: el('wh-edi-labels'),
    whBusinessParcels: el('wh-business-parcels'),
    whPrivateParcels: el('wh-private-parcels'),
    whPackaging: el('wh-packaging'),
    whWms: el('wh-wms'),
    wireView: el('wire-view'),
    wireCanvas: el('wire-canvas'),
    wireSelectedTitle: el('wire-selected-title'),
    wireSelectedDetail: el('wire-selected-detail'),
    wireSelectedLinks: el('wire-selected-links'),
    wireRelationForm: el('wire-relation-form'),
    wireFrom: el('wire-from'),
    wireTo: el('wire-to'),
    wireLabel: el('wire-label'),
    wireCustomList: el('wire-custom-list'),
    wireDetailMode: el('wire-detail-mode'),
    wireExpand: el('wire-expand'),
    wireResetLayout: el('wire-reset-layout'),
    sheetView: el('sheet-view'),
    grid: el('sheet-grid'),
    selectedAddress: el('selected-address'),
    formulaInput: el('formula-input'),
    saveCell: el('save-cell'),
    targetRange: el('target-range'),
    rangeFormula: el('range-formula'),
    applyRange: el('apply-range'),
    clearRange: el('clear-range'),
    dependencyPanel: el('dependency-panel'),
    dropZone: el('drop-zone'),
    toast: el('toast'),
    resetSettings: el('reset-settings'),
    freightSummary: el('freight-summary'),
    kitBar: el('kit-bar'),
    kitBreakdownTable: el('kit-breakdown-table'),
    consolidatedSearch: el('consolidated-search'),
    consolidatedFilter: el('consolidated-filter'),
    consolidatedCount: el('consolidated-count'),
    consolidatedActions: el('consolidated-actions'),
    dashboardView: el('dashboard-view'),
    dashboardSubtitle: el('dashboard-subtitle'),
    dashboardBadge: el('dashboard-badge'),
    dashboardKitBar: el('dashboard-kit-bar'),
    dashboardIncomplete: el('dashboard-incomplete'),
    dashboardKpis: el('dashboard-kpis'),
    vatSummary: el('vat-summary'),
    groupList: el('group-list'),
    groupForm: el('group-form'),
    groupName: el('group-name'),
    chartWaterfall: el('chart-waterfall'),
    waterfallTable: el('waterfall-table'),
    chartCumulative: el('chart-cumulative'),
    cumulativeTable: el('cumulative-table'),
    chartGroups: el('chart-groups'),
    groupTable: el('group-table'),
    dashboardItems: el('dashboard-items'),
    scenarioSummary: el('scenario-summary'),
    scenarioReset: el('scenario-reset'),
    chartTooltip: el('chart-tooltip'),
    flowOverview: el('flow-overview'),
    flowCustomerDestination: el('flow-customer-destination'),
    flowCustomerNote: el('flow-customer-note'),
    destinationMapTitle: el('destination-map-title'),
    flowPalletLabel: el('flow-pallet-label'),
    flowPalletRate: el('flow-pallet-rate'),
    flowWarehouseTotal: el('flow-warehouse-total'),
    flowWarehouseStats: el('flow-warehouse-stats'),
    flowPerItem: el('flow-per-item'),
    flowLock: el('flow-lock'),
    flowResetOverrides: el('flow-reset-overrides'),
    flowCustomerStats: el('flow-customer-stats'),
    sheetActions: el('sheet-actions'),
    customsPopover: el('customs-popover'),
    flowMap: el('flow-map'),
    mapWorld: el('map-world'),
    mapNorway: el('map-norway'),
    mapInfo: el('map-info'),
    flowLanes: el('flow-lanes'),
    makerList: el('maker-list'),
    makersPanel: el('makers-panel'),
    mapStage: el('map-stage'),
    mapOverlay: el('map-overlay'),
    mapBoxes: el('map-boxes'),
    mapLegend: el('map-legend'),
    flowBoxWarehouse: el('flow-box-warehouse'),
    flowBoxCustomer: el('flow-box-customer'),
    flowLegWarehouse: el('flow-leg-warehouse'),
    flowLegCustomer: el('flow-leg-customer'),
    flowNote: el('flow-note'),
    vatState: el('vat-state'),
    vatCountries: el('vat-countries'),
    vatRoutes: el('vat-routes'),
  };

  initialize();

  async function initialize() {
    purgeLegacyLocalBusinessData();
    loadUiPreferences();
    // Translate and render the panel state before any network wait so the page never flashes the wrong language.
    setLanguage(state.language, false);
    renderSidebarState();
    populateCurrencySelect(dom.sourceCurrency, state.config.sourceCurrency);
    populateCurrencySelect(dom.outputCurrency, state.config.outputCurrency);
    document.querySelectorAll('[data-currency-select]').forEach((select) => populateCurrencySelect(select, getPath(select.dataset.bind)));
    populateOriginSelects();
    populateCustomerSelects();
    syncBoundInputs();
    await loadPrivateWarehouseRateFile();
    syncWarehouseRateInputs();
    bindEvents();
    detectVault();
    observeLanguage();
    setLanguage(state.language, false);
    renderSaveState();
    if (!window.XLSX || !FormulaEngine || !Core) {
      setStatus('The spreadsheet libraries could not be loaded. Check the internet connection and reload.', 'error');
    }
  }

  function bindEvents() {
    window.addEventListener('beforeunload', (event) => {
      if (!state.projectDirty) return;
      event.preventDefault();
      event.returnValue = '';
    });

    window.addEventListener('popstate', () => applyHashView());
    window.addEventListener('hashchange', () => { if (!state.suppressHash) applyHashView(); });
    dom.undoButton.addEventListener('click', () => (state.redoStack.length ? redo() : undo()));
    document.addEventListener('keydown', (event) => {
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== 'z') return;
      const field = event.target;
      const typing = field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement || field?.isContentEditable;
      if (typing) return;
      event.preventDefault();
      if (event.shiftKey) redo();
      else undo();
    });
    dom.file.addEventListener('change', (event) => {
      const [file] = event.target.files;
      if (file) openWorkbookFile(file);
      event.target.value = '';
    });
    dom.vaultForm.addEventListener('submit', unlockVault);
    dom.addKitFile.addEventListener('change', async (event) => {
      const files = [...event.target.files];
      if (files.length) await addKitWorkbooks(files);
      event.target.value = '';
    });
    dom.languageSelect.addEventListener('click', (event) => {
      const button = event.target.closest('[data-language]');
      if (!button || button.dataset.language === state.language) return;
      setLanguage(button.dataset.language);
      saveUiPreferences();
      saveSettingsSoon();
    });
    // File inputs are opened through <label> buttons, so make those reachable from the keyboard too.
    document.querySelectorAll('label.button[for][role="button"]').forEach((label) => label.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      document.getElementById(label.htmlFor)?.click();
    }));

    [dom.sidebarToggle, dom.sidebarCollapse, dom.sidebarExpand].forEach((button) => button.addEventListener('click', () => {
      state.sidebarCollapsed = !state.sidebarCollapsed;
      saveUiPreferences();
      renderSidebarState();
    }));

    ['dragenter', 'dragover'].forEach((name) => dom.dropZone.addEventListener(name, (event) => {
      event.preventDefault();
      dom.dropZone.classList.add('dragging');
    }));
    ['dragleave', 'drop'].forEach((name) => dom.dropZone.addEventListener(name, (event) => {
      event.preventDefault();
      dom.dropZone.classList.remove('dragging');
    }));
    dom.dropZone.addEventListener('drop', (event) => {
      const [file] = event.dataTransfer.files;
      if (file) openWorkbookFile(file);
    });

    // One delegated handler keeps every copy of a value (sidebar, flow, dashboard) in sync.
    document.addEventListener('input', handleBoundEvent);
    document.addEventListener('change', handleBoundEvent);
    document.addEventListener('focusout', (event) => {
      if (event.target?.matches?.('[data-bind], [data-kit-bind], [data-group-field], [data-override]')) setTimeout(syncBoundInputs, 0);
    });

    dom.sourceCurrency.addEventListener('change', () => {
      state.config.sourceCurrency = dom.sourceCurrency.value;
      saveSettingsSoon();
      rebuildConsolidation();
      refreshRates(false);
      renderMappings();
    });
    dom.outputCurrency.addEventListener('change', () => {
      state.config.outputCurrency = dom.outputCurrency.value;
      state.rates = {};
      saveSettingsSoon();
      rebuildConsolidation();
      renderMappings();
      refreshRates(false);
    });
    dom.resetSettings.addEventListener('click', resetProjectSettings);
    dom.consolidatedSearch.addEventListener('input', () => {
      state.consolidatedSearch = dom.consolidatedSearch.value;
      renderConsolidated();
    });
    dom.consolidatedFilter.addEventListener('change', () => {
      state.consolidatedFilter = dom.consolidatedFilter.value;
      renderConsolidated();
    });
    dom.consolidatedTable.addEventListener('click', handleConsolidatedTableClick);
    dom.consolidatedTable.addEventListener('keydown', handleConsolidatedTableKeydown);
    // The Columns menu closes on an outside click or Escape, so it never covers the table.
    document.addEventListener('click', (event) => {
      if (dom.columnMenu.open && !dom.columnMenu.contains(event.target)) dom.columnMenu.open = false;
    });
    dom.columnMenu.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') dom.columnMenu.open = false;
    });
    dom.columnsDone.addEventListener('click', () => { dom.columnMenu.open = false; });
    dom.columnsAll.addEventListener('click', () => {
      state.hiddenColumns = [];
      saveUiPreferences();
      renderConsolidated();
    });
    dom.columnsReset.addEventListener('click', () => {
      state.hiddenColumns = [...DEFAULT_HIDDEN_COLUMNS];
      saveUiPreferences();
      renderConsolidated();
    });
    dom.columnList.addEventListener('change', (event) => {
      const box = event.target.closest('[data-column-key]');
      if (!box) return;
      toggleColumn(box.dataset.columnKey, box.checked);
    });
    dom.groupForm.addEventListener('submit', (event) => {
      event.preventDefault();
      const group = addGroup(dom.groupName.value);
      if (group) dom.groupName.value = '';
    });
    dom.scenarioReset.addEventListener('click', () => {
      state.scenario = { salesDelta: 0, purchaseDelta: 0, inboundDelta: 0, warehouseDelta: 0 };
      saveSettingsSoon();
      refreshViews();
    });
    dom.groupList.addEventListener('click', (event) => {
      const remove = event.target.closest('[data-remove-group]');
      if (remove) deleteGroup(remove.dataset.removeGroup);
    });
    document.addEventListener('click', (event) => {
      if (!event.target.closest?.('[data-vat-lock]')) return;
      state.vat.unlocked = !state.vat.unlocked;
      renderVatControls();
      if (state.vat.unlocked) document.getElementById('vat-input')?.focus();
    });
    dom.flowLock.addEventListener('click', () => {
      state.flowUnlocked = !state.flowUnlocked;
      renderFlowLock();
      if (state.flowUnlocked) dom.flowPerItem.querySelector('[data-override]')?.focus();
    });
    dom.flowResetOverrides.addEventListener('click', () => {
      FLOW_OVERRIDE_KEYS.forEach((key) => { state.flowOverrides[key] = null; });
      saveSettingsSoon();
      refreshViews();
      showToast('Per-item values reset to the calculated amounts.');
    });
    [dom.kitBar, dom.dashboardKitBar].forEach((bar) => bar.addEventListener('click', (event) => {
      const only = event.target.closest('[data-kit-only]');
      if (only) setKitsOnly(only.dataset.kitOnly);
      const chip = event.target.closest('[data-kit-toggle]');
      if (chip) setKitEnabled(chip.dataset.kitToggle, chip.dataset.kitEnable === 'true');
    }));
    [dom.consolidatedActions, dom.sheetActions].forEach((bar) => {
      bar.addEventListener('change', handleSelectionBarChange);
      bar.addEventListener('click', handleSelectionBarClick);
      bar.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' && event.target.matches('input')) {
          event.preventDefault();
          applySelectionAction(bar);
        }
      });
    });
    bindChartTooltip();
    bindCustomsPopover();
    bindMapEvents();
    if ('ResizeObserver' in window) {
      let resizeFrame = 0;
      let lastWidths = '';
      const observer = new ResizeObserver(() => {
        cancelAnimationFrame(resizeFrame);
        resizeFrame = requestAnimationFrame(() => {
          const widths = [dom.chartWaterfall, dom.chartCumulative, dom.chartGroups, dom.mapStage].map((node) => node.clientWidth).join(',');
          if (widths === lastWidths) return;
          lastWidths = widths;
          if (state.activeView === VIEW.dashboard) renderDashboardCharts();
          if (state.activeView === VIEW.wire && state.flowMode === 'map') renderMaps();
        });
      });
      [dom.chartWaterfall, dom.chartCumulative, dom.chartGroups, dom.mapStage].forEach((node) => observer.observe(node));
    }
    dom.refreshRates.addEventListener('click', () => refreshRates(true));
    dom.exportButton.addEventListener('click', exportWorkbook);
    dom.guideToggle.addEventListener('click', () => toggleSetupGuide());
    dom.guideClose.addEventListener('click', () => toggleSetupGuide(false));
    dom.setupChecklist.addEventListener('click', handleSetupGuideAction);

    dom.saveCell.addEventListener('click', saveSelectedCell);
    dom.formulaInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') saveSelectedCell();
    });
    dom.applyRange.addEventListener('click', () => applyToRange(false));
    dom.clearRange.addEventListener('click', () => applyToRange(true));
    dom.variableForm.addEventListener('submit', addVariable);
    dom.secureExport.addEventListener('click', exportSecureProject);
    dom.secureFile.addEventListener('change', (event) => {
      const [file] = event.target.files;
      if (file) importSecureProject(file);
      event.target.value = '';
    });
    dom.secureDownload.addEventListener('click', () => markProjectSaved('Encrypted project downloaded'));
    dom.wireRelationForm.addEventListener('submit', addWireRelation);
    dom.wireDetailMode.addEventListener('change', () => {
      state.wireMode = dom.wireDetailMode.value === 'underhood' ? 'underhood' : 'overview';
      state.selectedWireNode = state.wireMode === 'underhood' ? 'column.quantity' : 'consolidated.quantity';
      renderWireView();
    });
    dom.wireExpand.addEventListener('click', toggleWireExpanded);
    dom.wireResetLayout.addEventListener('click', resetWireLayout);
    dom.wireCanvas.addEventListener('pointerdown', startWireDrag);
    dom.wireCanvas.addEventListener('pointermove', moveWireDrag);
    dom.wireCanvas.addEventListener('pointerup', finishWireDrag);
    dom.wireCanvas.addEventListener('pointercancel', finishWireDrag);
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && state.wireExpanded) toggleWireExpanded(false);
    });
    bindWarehouseEvents();
  }

  function toggleWireExpanded(force) {
    state.wireExpanded = typeof force === 'boolean' ? force : !state.wireExpanded;
    dom.wireView.classList.toggle('wire-expanded', state.wireExpanded);
    dom.wireExpand.textContent = state.wireExpanded ? 'Close expanded view' : 'Expand view';
    dom.wireExpand.setAttribute('aria-pressed', String(state.wireExpanded));
  }

  function bindWarehouseEvents() {
    const numericFields = [
      [dom.whShelfShare, 'shelfShare', 0],
      [dom.whDrawerShare, 'drawerShare', 0],
      [dom.whBinsPerShelf, 'binsPerShelf', 1],
      [dom.whShelvesPerRack, 'shelvesPerRack', 1],
      [dom.whUnitsPerBin, 'unitsPerBin', 1],
      [dom.whUnitsDrawer, 'unitsPerDrawer', 1],
      [dom.whUnitsPallet, 'unitsPerPallet', 1],
      [dom.whReceipts, 'receipts', 0],
      [dom.whReceiptLines, 'receiptLines', 0],
      [dom.whOrders, 'orders', 0],
      [dom.whOrderLines, 'orderLines', 0],
      [dom.whEdiLabels, 'ediLabels', 0],
      [dom.whBusinessParcels, 'businessParcels', 0],
      [dom.whPrivateParcels, 'privateParcels', 0],
      [dom.whPackaging, 'packaging', 0],
    ];

    numericFields.forEach(([input, key, minimum]) => {
      input.addEventListener('input', () => {
        let value = Math.max(minimum, toNumber(input.value) ?? minimum);
        if (key === 'shelfShare' || key === 'drawerShare') value = clamp(value, 0, 100);
        if (key === 'binsPerShelf') value = clamp(Math.round(value), 1, 12);
        if (key === 'shelvesPerRack') value = clamp(Math.round(value), 1, 8);
        state.warehouse[key] = value;
        state.warehouse.unitsPerShelf = state.warehouse.binsPerShelf * state.warehouse.unitsPerBin;
        dom.whUnitsShelf.value = String(state.warehouse.unitsPerShelf);
        if (key === 'shelfShare' || key === 'drawerShare') constrainStorageShares(key);
        saveSettingsSoon();
        refreshViews();
      });
    });

    dom.whPalletType.addEventListener('change', () => {
      state.warehouse.palletType = dom.whPalletType.value === 'sea' ? 'sea' : 'eu';
      saveSettingsSoon();
      refreshViews();
    });
    dom.whPalletHeight.addEventListener('change', () => {
      state.warehouse.palletHeight = dom.whPalletHeight.value === '220' ? '220' : '120';
      saveSettingsSoon();
      refreshViews();
    });
    // The WMS checkbox and quoted-rate inputs use data-bind, shared with the wire-view flow.

    dom.warehouseRatesExport.addEventListener('click', exportWarehouseRates);
    dom.warehouseRatesFile.addEventListener('change', (event) => {
      const [file] = event.target.files;
      if (file) importWarehouseRates(file);
      event.target.value = '';
    });

    dom.warehousePreview.addEventListener('wheel', (event) => {
      event.preventDefault();
      const direction = event.deltaY < 0 ? 1 : -1;
      zoomWarehousePreview(direction * 0.12);
    }, { passive: false });
    dom.warehousePreview.addEventListener('click', (event) => {
      const button = event.target.closest('[data-preview-action]');
      if (!button) return;
      const action = button.dataset.previewAction;
      if (action === 'in') zoomWarehousePreview(0.18);
      if (action === 'out') zoomWarehousePreview(-0.18);
      if (action === 'reset') resetWarehousePreview();
    });
    dom.warehousePreview.addEventListener('pointerdown', (event) => {
      if (event.target.closest('button')) return;
      dom.warehousePreview.setPointerCapture(event.pointerId);
      state.warehousePreviewDrag = { pointerId: event.pointerId, x: event.clientX, y: event.clientY };
      dom.warehousePreview.classList.add('dragging');
    });
    dom.warehousePreview.addEventListener('pointermove', (event) => {
      const drag = state.warehousePreviewDrag;
      if (!drag || drag.pointerId !== event.pointerId) return;
      const svg = dom.warehousePreview.querySelector('svg');
      const rect = svg?.getBoundingClientRect();
      if (!rect?.width || !rect?.height) return;
      state.warehousePreviewTransform.x += (event.clientX - drag.x) * (760 / rect.width);
      state.warehousePreviewTransform.y += (event.clientY - drag.y) * (420 / rect.height);
      drag.x = event.clientX;
      drag.y = event.clientY;
      applyWarehousePreviewTransform();
    });
    const endPreviewDrag = (event) => {
      if (!state.warehousePreviewDrag || state.warehousePreviewDrag.pointerId !== event.pointerId) return;
      state.warehousePreviewDrag = null;
      dom.warehousePreview.classList.remove('dragging');
    };
    dom.warehousePreview.addEventListener('pointerup', endPreviewDrag);
    dom.warehousePreview.addEventListener('pointercancel', endPreviewDrag);
  }

  function constrainStorageShares(changedKey) {
    const warehouse = state.warehouse;
    const otherKey = changedKey === 'shelfShare' ? 'drawerShare' : 'shelfShare';
    if (warehouse.shelfShare + warehouse.drawerShare <= 100) {
      warehouse.shareMessage = '';
      return;
    }
    warehouse[otherKey] = Math.max(0, 100 - warehouse[changedKey]);
    const otherInput = otherKey === 'shelfShare' ? dom.whShelfShare : dom.whDrawerShare;
    otherInput.value = String(warehouse[otherKey]);
    warehouse.shareMessage = 'Storage shares were capped at 100%; the remaining share was adjusted automatically.';
  }

  /* ---------- Two-way value binding ---------- */

  function getPath(path) {
    return String(path || '').split('.').reduce((target, key) => (target === null || target === undefined ? undefined : target[key]), state);
  }

  function setPath(path, value) {
    const keys = String(path).split('.');
    const last = keys.pop();
    const target = keys.reduce((object, key) => object?.[key], state);
    if (target && typeof target === 'object') target[last] = value;
  }

  function readNumberInput(input) {
    const value = toNumber(input.value);
    if (value === null) return undefined;
    const min = input.min !== '' ? Number(input.min) : -Infinity;
    const max = input.max !== '' ? Number(input.max) : Infinity;
    return clamp(value, min, max) / Number(input.dataset.scale || 1);
  }

  function writeInputValue(input, value) {
    if (input.type === 'checkbox') {
      input.checked = Boolean(value);
    } else if (input.type === 'number') {
      const scaled = Number.isFinite(value) ? value * Number(input.dataset.scale || 1) : null;
      input.value = scaled === null ? '' : String(Number(scaled.toFixed(6)));
    } else if (input.value !== String(value ?? '')) {
      input.value = value ?? '';
    }
  }

  function handleBoundEvent(event) {
    const input = event.target;
    if (!(input instanceof HTMLInputElement || input instanceof HTMLSelectElement)) return;
    const discrete = input.tagName === 'SELECT' || input.type === 'checkbox';
    if ((event.type === 'input' && discrete) || (event.type === 'change' && !discrete && !input.dataset.makerBind)) return;

    if (input.dataset.bind) {
      const value = input.type === 'checkbox' ? input.checked : input.type === 'number' ? readNumberInput(input) : input.value;
      if (value === undefined) return;
      setPath(input.dataset.bind, value);
      if (discrete) onBoundChange(input.dataset.bind);
      else {
        onBoundChange(input.dataset.bind, { defer: true });
        scheduleRecalc(debouncedKindFor(input.dataset.bind));
      }
    } else if (input.dataset.makerBind) {
      handleMakerInput(input, event.type);
    } else if (input.dataset.kitBind) {
      handleKitInput(input);
    } else if (input.dataset.groupField) {
      handleGroupInput(input);
    } else if (input.dataset.vatCountry) {
      state.vat.byCountry[input.dataset.vatCountry] = input.value === 'price' ? 'price' : 'none';
      saveSettingsSoon();
      rebuildConsolidation();
    } else if (input.dataset.override) {
      if (input.readOnly) return;
      const value = input.value.trim() === '' ? null : readNumberInput(input);
      if (value === undefined) return;
      state.flowOverrides[input.dataset.override] = value;
      saveSettingsSoon();
      scheduleRecalc('views');
    }
  }

  // Text fields that only affect the rendered views, rather than the consolidation itself.
  function debouncedKindFor(bind) {
    return bind.startsWith('scenario.') ? 'views' : 'consolidation';
  }

  /* ---------- Debounced recalculation ---------- */

  // Typing in a number field fires an event per keystroke. Recomputing and re-rendering the whole
  // table each time is what made the form feel sticky, so the heavy work is coalesced into one call
  // shortly after typing stops. Dropdown and checkbox changes are discrete and still run at once.
  function scheduleRecalc(kind) {
    if (state.recalcTimer) clearTimeout(state.recalcTimer);
    state.pendingRecalc = kind;
    state.recalcTimer = setTimeout(() => {
      state.recalcTimer = null;
      flushRecalc();
    }, RECALC_DEBOUNCE_MS);
  }

  function flushRecalc() {
    if (state.recalcTimer) {
      clearTimeout(state.recalcTimer);
      state.recalcTimer = null;
    }
    const kind = state.pendingRecalc;
    state.pendingRecalc = null;
    if (!kind) return;
    if (kind === 'consolidation') {
      rebuildConsolidation();
      if (state.activeView === VIEW.consolidated) renderConsolidated();
    } else {
      refreshViews();
    }
  }

  function onBoundChange(path, options = {}) {
    if (path.startsWith('warehouse.rates.')) persistWarehouseRates();
    if (path.startsWith('warehouse.')) state.warehouse.unitsPerShelf = state.warehouse.binsPerShelf * state.warehouse.unitsPerBin;
    if (path === 'warehouse.plannedPallets') state.warehouse.plannedPallets = Math.max(0, Math.round(state.warehouse.plannedPallets || 0));
    saveSettingsSoon();
    // A deferred change is a text field mid-edit: state is already updated, but the caller schedules
    // one coalesced rebuild instead of paying for a full render on every keystroke.
    if (options.defer) return;
    if (/Currency$/.test(path)) {
      rebuildConsolidation();
      refreshRates(false);
      return;
    }
    if (path === 'customer.country') {
      state.customer.country = DESTINATIONS[state.customer.country] ? state.customer.country : '578';
      rebuildConsolidation();
      return;
    }
    if (path === 'supplier.country') {
      setSupplierOrigin(state.supplier.country);
      return;
    }
    if (path.startsWith('scenario.')) {
      refreshViews();
      return;
    }
    if (path.startsWith('warehouse.') || path.startsWith('supplier.') || path.startsWith('sales.')) {
      if (path === 'warehouse.country') renderMappings();
      if (path === 'sales.defaultValue' || path === 'sales.defaultMode') rebuildConsolidation();
      else refreshViews();
      return;
    }
    rebuildConsolidation();
  }

  function syncBoundInputs() {
    const active = document.activeElement;
    document.querySelectorAll('[data-bind]').forEach((input) => {
      if (input !== active) writeInputValue(input, getPath(input.dataset.bind));
    });
    document.querySelectorAll('[data-kit-bind]').forEach((input) => {
      const kit = kitSettings(input.dataset.kitBind);
      if (input !== active) writeInputValue(input, kit[input.dataset.kitField]);
      if (input.dataset.kitField === 'discountRate') input.placeholder = `global ${formatPercent(state.config.discount)}`;
      if (input.dataset.kitField === 'dutyRate') input.placeholder = `global ${formatPercent(state.freight.dutyRate)}`;
      if (input.dataset.kitField === 'shippingAmount') {
        input.disabled = kit.shippingMode === 'sheet';
        input.placeholder = kit.shippingMode === 'sheet' ? 'sheet' : '0';
      }
      if (input.dataset.kitField === 'shippingCurrency') input.hidden = kit.shippingMode === 'percent';
    });
    dom.flowNote.textContent = `Edit values on the route; every view recalculates. Amounts in ${state.config.outputCurrency} unless marked.`;
    syncMakerInputs();
    syncControlsFromState(active);
  }

  /* ---------- Kits (source sheets) ---------- */

  // A new kit starts as a kit from the project's supplier (TEI Rock Drills, USA by default). Discount, source
  // currency and duty stay empty, which means "use the project-wide value", so changing the sidebar still works.
  function newKitSettings(sheetName) {
    return {
      ...KIT_DEFAULTS,
      manufacturer: state.supplier.name || sheetName,
      originCountry: state.supplier.country || KIT_DEFAULTS.originCountry,
      shippingCurrency: state.config.outputCurrency,
    };
  }

  function kitDiscount(kit) {
    return Number.isFinite(kit.discountRate) ? kit.discountRate : state.config.discount;
  }

  function kitCurrency(kit) {
    return kit.defaultCurrency || state.config.sourceCurrency;
  }

  // Added on top of the live exchange rate (in the output currency per unit of the source currency).
  function kitFxMargin(kit) {
    return Number.isFinite(kit.fxMargin) && kit.fxMargin > 0 ? kit.fxMargin : 0;
  }

  function sanitizeMaker(maker) {
    return {
      location: typeof maker.location === 'string' ? maker.location.slice(0, 80) : '',
      shipmentMode: maker.shipmentMode === 'own' ? 'own' : 'project',
      shipmentAmount: finiteOr(maker.shipmentAmount, 0, 0),
      shipmentCurrency: CURRENCIES.includes(maker.shipmentCurrency) ? maker.shipmentCurrency : MAKER_DEFAULTS.shipmentCurrency,
      clearanceAmount: finiteOr(maker.clearanceAmount, 0, 0),
      clearanceCurrency: CURRENCIES.includes(maker.clearanceCurrency) ? maker.clearanceCurrency : MAKER_DEFAULTS.clearanceCurrency,
      insuranceRate: nullableFraction(maker.insuranceRate),
    };
  }

  function makerSettings(name) {
    if (!state.makers[name]) {
      state.makers[name] = { ...MAKER_DEFAULTS, shipmentCurrency: state.freight.consolidatedCurrency, clearanceCurrency: state.freight.clearanceCurrency };
    }
    return state.makers[name];
  }

  const makerHasOwnShipment = (name) => state.makers[name]?.shipmentMode === 'own';

  function makerInsuranceRate(name) {
    const own = state.makers[name]?.insuranceRate;
    return Number.isFinite(own) ? own : state.freight.insuranceRate;
  }

  // Every manufacturer that has at least one kit sheet, in sheet order.
  function makerNames() {
    if (!state.workbook) return [];
    return [...new Set(state.workbook.SheetNames.filter(isSourceKit).map((name) => kitSettings(name).manufacturer || name))];
  }

  function makerKits(name) {
    return state.workbook ? state.workbook.SheetNames.filter(isSourceKit).filter((sheet) => (kitSettings(sheet).manufacturer || sheet) === name) : [];
  }

  const TREATMENT_LABELS = { domestic: 'Domestic', intraeu: 'Intra-EU (no duty)', import: 'Import (customs clearance)' };

  function kitRouteSummary(kit) {
    const treatment = Core.customsTreatment(kit.originCountry, state.warehouse.country);
    return [`${originName(kit.originCountry)} → ${destinationConfig().name}`, TREATMENT_LABELS[treatment], `${kitCurrency(kit)} source prices`].join(' · ');
  }

  function updateKitRouteSummaries() {
    document.querySelectorAll('[data-kit-route]').forEach((node) => { node.textContent = kitRouteSummary(kitSettings(node.dataset.kitRoute)); });
  }

  function kitSettings(sheetName) {
    if (!state.kits[sheetName]) state.kits[sheetName] = newKitSettings(sheetName);
    return state.kits[sheetName];
  }

  function handleKitInput(input) {
    const kit = kitSettings(input.dataset.kitBind);
    const field = input.dataset.kitField;
    const parsed = readKitField(field, input, input.dataset.kitBind);
    if (!parsed) return;
    kit[field] = parsed.value;
    saveSettingsSoon();
    // Free-text kit fields (manufacturer) are typed, so the rebuild is coalesced like other text inputs.
    const discrete = input.tagName === 'SELECT' || input.type === 'checkbox';
    if (discrete) {
      rebuildConsolidation();
      if (['shippingCurrency', 'defaultCurrency'].includes(field)) refreshRates(false);
      if (field === 'originCountry') renderMaps();
      updateKitRouteSummaries();
    } else {
      scheduleRecalc('consolidation');
    }
  }

  function isSourceKit(sheetName) {
    const mapping = state.mappings[sheetName];
    return Boolean(mapping && mapping.part !== null && mapping.quantity !== null && !isGeneratedSheet(state.workbook?.Sheets[sheetName]));
  }

  // One click to look at a single manufacturer's kits (or all of them again); undo restores the previous choice.
  function setKitsOnly(maker) {
    const names = state.workbook.SheetNames.filter(isSourceKit);
    pushHistory(maker ? `Only ${maker} kits` : 'All kits', () => {
      names.forEach((name) => { state.mappings[name].enabled = !maker || (kitSettings(name).manufacturer || name) === maker; });
      renderMappings();
      saveSettingsSoon();
      rebuildConsolidation();
      refreshRates(false);
    });
  }

  function setKitEnabled(sheetName, enabled) {
    const mapping = state.mappings[sheetName];
    if (!mapping) return;
    pushHistory(`${sheetName} ${enabled ? 'added to' : 'removed from'} the consolidation`, () => {
      mapping.enabled = enabled;
      renderMappings();
      saveSettingsSoon();
      rebuildConsolidation();
      refreshRates(false);
    });
    showToast(enabled
      ? `${sheetName} added to the consolidation. Everything was recalculated. Ctrl+Z reverts it.`
      : `${sheetName} removed from the consolidation. Everything was recalculated. Ctrl+Z reverts it.`);
  }

  function createKitFreightControl(sheetName) {
    const kit = kitSettings(sheetName);
    const wrapper = document.createElement('div');
    wrapper.className = 'kit-freight-control kit-profile-control';
    const profileCaption = document.createElement('span');
    profileCaption.className = 'kit-freight-caption';
    profileCaption.textContent = 'Kit profile & route';
    const manufacturer = document.createElement('input');
    manufacturer.dataset.kitBind = sheetName;
    manufacturer.dataset.kitField = 'manufacturer';
    manufacturer.value = kit.manufacturer || sheetName;
    manufacturer.placeholder = 'Manufacturer';
    manufacturer.setAttribute('aria-label', `${sheetName} manufacturer`);
    const origin = document.createElement('select');
    origin.dataset.kitBind = sheetName;
    origin.dataset.kitField = 'originCountry';
    origin.setAttribute('aria-label', `${sheetName} origin country`);
    populateOriginSelect(origin, kit.originCountry);
    const sourceCurrency = document.createElement('select');
    sourceCurrency.dataset.kitBind = sheetName;
    sourceCurrency.dataset.kitField = 'defaultCurrency';
    sourceCurrency.setAttribute('aria-label', `${sheetName} default source currency`);
    sourceCurrency.append(new Option('Default', ''), ...CURRENCIES.map((code) => new Option(code, code)));
    sourceCurrency.value = kit.defaultCurrency || '';
    const percentField = (field, caption) => {
      const label = document.createElement('label');
      label.className = 'kit-discount-field';
      label.textContent = caption;
      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.max = '100';
      input.step = '0.1';
      input.dataset.scale = '100';
      input.dataset.kitBind = sheetName;
      input.dataset.kitField = field;
      writeInputValue(input, kit[field]);
      label.append(input);
      return label;
    };
    const discount = percentField('discountRate', 'Discount %');
    const duty = percentField('dutyRate', 'Duty %');
    const buffer = document.createElement('label');
    buffer.className = 'kit-discount-field';
    buffer.textContent = 'Rate buffer';
    const bufferInput = document.createElement('input');
    bufferInput.type = 'number';
    bufferInput.min = '0';
    bufferInput.step = '0.01';
    bufferInput.dataset.kitBind = sheetName;
    bufferInput.dataset.kitField = 'fxMargin';
    writeInputValue(bufferInput, kit.fxMargin);
    buffer.append(bufferInput);
    const caption = document.createElement('span');
    caption.className = 'kit-freight-caption';
    caption.textContent = 'Kit freight';
    const mode = document.createElement('select');
    mode.dataset.kitBind = sheetName;
    mode.dataset.kitField = 'shippingMode';
    mode.setAttribute('aria-label', `${sheetName} freight mode`);
    Object.entries(KIT_SHIPPING_MODES).forEach(([value, label]) => mode.append(new Option(label, value, false, kit.shippingMode === value)));
    const amount = document.createElement('input');
    amount.type = 'number';
    amount.min = '0';
    amount.step = '0.01';
    amount.dataset.kitBind = sheetName;
    amount.dataset.kitField = 'shippingAmount';
    amount.setAttribute('aria-label', `${sheetName} freight amount`);
    writeInputValue(amount, kit.shippingAmount);
    amount.disabled = kit.shippingMode === 'sheet';
    const currency = document.createElement('select');
    currency.dataset.kitBind = sheetName;
    currency.dataset.kitField = 'shippingCurrency';
    currency.setAttribute('aria-label', `${sheetName} freight currency`);
    populateCurrencySelect(currency, kit.shippingCurrency);
    const help = document.createElement('small');
    help.textContent = 'Sheet values use the Frakt column per line. Per line applies one amount to every line; Whole kit splits one amount across the kit’s stocked lines.';
    const route = document.createElement('small');
    route.className = 'kit-route-summary';
    route.dataset.kitRoute = sheetName;
    route.textContent = kitRouteSummary(kit);
    wrapper.append(profileCaption, manufacturer, origin, sourceCurrency, discount, duty, buffer, caption, mode, amount, currency, help, route);
    return wrapper;
  }

  /* ---------- Manufacturers: one route each, one warehouse for all ---------- */

  // Kit fields a manufacturer owns: changing one on the manufacturer changes it on every kit of that manufacturer.
  const MAKER_KIT_FIELDS = new Set(['originCountry', 'defaultCurrency', 'discountRate', 'dutyRate', 'fxMargin', 'shippingMode', 'shippingAmount', 'shippingCurrency']);

  // Reads one kit field from an input; null when the input holds nothing usable yet.
  function readKitField(field, input, fallbackName = '') {
    switch (field) {
      case 'shippingAmount': {
        const value = readNumberInput(input);
        return value === undefined ? null : { value: Math.max(0, value) };
      }
      case 'shippingMode': return { value: KIT_SHIPPING_MODES[input.value] ? input.value : 'sheet' };
      case 'shippingCurrency': return { value: CURRENCIES.includes(input.value) ? input.value : KIT_DEFAULTS.shippingCurrency };
      case 'manufacturer': return { value: String(input.value || '').trim().slice(0, 80) || fallbackName };
      case 'originCountry': return { value: Core.normalizeCountryCode(input.value, KIT_DEFAULTS.originCountry) };
      case 'defaultCurrency': return { value: CURRENCIES.includes(input.value) ? input.value : null };
      case 'fxMargin': {
        const value = readNumberInput(input);
        return value === undefined ? null : { value: Math.max(0, value) };
      }
      case 'discountRate':
      case 'dutyRate': {
        if (input.value.trim() === '') return { value: null };
        const value = readNumberInput(input);
        return value === undefined ? null : { value: clamp(value, 0, 1) };
      }
      default: return null;
    }
  }

  function makerLocation(name) {
    const own = state.makers[name]?.location;
    if (own) return own;
    const sheet = makerKits(name)[0];
    const origin = sheet ? kitSettings(sheet).originCountry : state.supplier.country;
    return ORIGIN_PRESETS[origin]?.place || originName(origin);
  }

  function makerFieldValue(name, field) {
    if (field === 'name') return name;
    if (field === 'location') return makerLocation(name);
    if (MAKER_KIT_FIELDS.has(field)) {
      const sheet = makerKits(name)[0];
      return sheet ? kitSettings(sheet)[field] : null;
    }
    return makerSettings(name)[field];
  }

  function makerOrigin(name) {
    return makerFieldValue(name, 'originCountry') || state.supplier.country;
  }

  // "EUR → SEK: 11.70 live rate + 0.40 buffer = 12.10", the number the prices are actually converted with.
  function makerRateText(name) {
    const sheet = makerKits(name)[0];
    const kit = sheet ? kitSettings(sheet) : null;
    const currency = kit ? kitCurrency(kit) : state.config.sourceCurrency;
    const output = state.config.outputCurrency;
    if (currency === output) return `Prices are already in ${output}, so no conversion is needed.`;
    const live = fxToOutput(currency);
    if (live === null) return `${currency} → ${output}: the live rate has not loaded yet.`;
    const margin = kit ? kitFxMargin(kit) : 0;
    return `${currency} → ${output}: ${formatNumber(live, 4)} live rate + ${formatNumber(margin, 2)} buffer = ${formatNumber(live + margin, 4)}`;
  }

  // One labelled input bound to a manufacturer's setting; the same builder serves the sidebar and the route view.
  function makerControl(name, field, spec = {}) {
    const label = document.createElement('label');
    label.className = spec.className || 'maker-field';
    const caption = document.createElement('span');
    caption.textContent = spec.caption;
    let control;
    if (spec.kind === 'select' || spec.kind === 'country' || spec.kind === 'currency') {
      control = document.createElement('select');
      if (spec.kind === 'select') (spec.options || []).forEach(([value, text]) => control.append(new Option(text, value)));
      if (spec.kind === 'currency') {
        if (spec.allowDefault) control.append(new Option('Default', ''));
        CURRENCIES.forEach((code) => control.append(new Option(code, code)));
      }
    } else {
      control = document.createElement('input');
      control.type = spec.kind === 'text' ? 'text' : 'number';
      if (spec.kind === 'text') {
        control.maxLength = 80;
        control.spellcheck = false;
      } else {
        control.min = '0';
        if (spec.max !== undefined) control.max = String(spec.max);
        control.step = spec.step || 'any';
        if (spec.scale) control.dataset.scale = String(spec.scale);
      }
    }
    control.dataset.makerBind = name;
    control.dataset.makerField = field;
    control.setAttribute('aria-label', `${name} ${spec.caption}`);
    if (spec.suffix) {
      const wrap = document.createElement('span');
      wrap.className = 'input-suffix';
      const unit = document.createElement('span');
      unit.textContent = spec.suffix;
      wrap.append(control, unit);
      label.append(caption, wrap);
    } else {
      label.append(caption, control);
    }
    if (spec.country) populateOriginSelect(control, makerFieldValue(name, field));
    return label;
  }

  const MAKER_SHIPMENT_OPTIONS = [['project', 'Shares the project shipment'], ['own', 'Ships on its own']];
  const KIT_MODE_OPTIONS = () => Object.entries(KIT_SHIPPING_MODES);

  // Freight add-on for the manufacturer's kits: a mode, an amount, and the unit that fits the mode.
  function makerFreightControls(name) {
    const mode = makerControl(name, 'shippingMode', { kind: 'select', caption: 'Kit freight', options: KIT_MODE_OPTIONS() });
    const amount = makerControl(name, 'shippingAmount', { kind: 'number', caption: 'Amount', step: '0.01' });
    const currency = makerControl(name, 'shippingCurrency', { kind: 'currency', caption: 'Currency' });
    currency.dataset.makerShow = 'amount';
    const percent = document.createElement('small');
    percent.className = 'maker-unit';
    percent.dataset.makerShow = 'percent';
    percent.textContent = '% of the item value';
    return [mode, amount, currency, percent];
  }

  function buildMakerCard(name) {
    const card = document.createElement('article');
    card.className = 'maker-card';
    card.dataset.maker = name;
    const head = document.createElement('div');
    head.className = 'maker-card-head';
    const title = makerControl(name, 'name', { kind: 'text', caption: 'Manufacturer', className: 'maker-field maker-name' });
    const count = document.createElement('small');
    count.className = 'maker-count';
    count.dataset.makerCount = name;
    head.append(title, count);
    const grid = document.createElement('div');
    grid.className = 'maker-grid';
    const rate = document.createElement('small');
    rate.className = 'maker-rate';
    rate.dataset.makerRate = name;
    const route = document.createElement('small');
    route.className = 'maker-route';
    route.dataset.makerRoute = name;
    grid.append(
      makerControl(name, 'originCountry', { kind: 'country', country: true, caption: 'Country of origin' }),
      makerControl(name, 'defaultCurrency', { kind: 'currency', allowDefault: true, caption: 'Price currency' }),
      makerControl(name, 'discountRate', { kind: 'number', caption: 'Discount', suffix: '%', scale: 100, max: 100, step: '0.1' }),
      makerControl(name, 'dutyRate', { kind: 'number', caption: 'Customs duty', suffix: '%', scale: 100, max: 100, step: '0.1' }),
      makerControl(name, 'fxMargin', { kind: 'number', caption: 'Rate buffer (added to the live rate)', step: '0.01', className: 'maker-field span-2' }),
      rate,
      ...makerFreightControls(name),
      makerControl(name, 'shipmentMode', { kind: 'select', caption: 'Consolidated shipment', options: MAKER_SHIPMENT_OPTIONS, className: 'maker-field span-2' }),
    );
    if (makerHasOwnShipment(name)) {
      grid.append(
        makerControl(name, 'shipmentAmount', { kind: 'number', caption: 'Shipment', step: '1' }),
        makerControl(name, 'shipmentCurrency', { kind: 'currency', caption: 'Currency' }),
        makerControl(name, 'clearanceAmount', { kind: 'number', caption: 'Clearance & broker', step: '1' }),
        makerControl(name, 'clearanceCurrency', { kind: 'currency', caption: 'Currency' }),
        makerControl(name, 'insuranceRate', { kind: 'number', caption: 'Insurance', suffix: '%', scale: 100, max: 100, step: '0.1' }),
      );
    }
    card.append(head, grid, route);
    return card;
  }

  function renderMakers() {
    if (!dom.makerList) return;
    const names = state.workbook ? makerNames() : [];
    const signature = names.map((name) => `${name}|${makerKits(name).length}|${makerSettings(name).shipmentMode}`).join('\u0001');
    if (dom.makerList.dataset.signature !== signature) {
      dom.makerList.dataset.signature = signature;
      dom.makerList.replaceChildren(...names.map(buildMakerCard));
    }
    dom.makersPanel.classList.toggle('hidden', !names.length);
    syncMakerInputs();
  }

  function syncMakerInputs() {
    const active = document.activeElement;
    document.querySelectorAll('[data-maker-bind]').forEach((input) => {
      const name = input.dataset.makerBind;
      const field = input.dataset.makerField;
      if (input !== active) {
        const value = makerFieldValue(name, field);
        if (input.tagName === 'SELECT') {
          if (field === 'originCountry') populateOriginSelect(input, value || state.supplier.country);
          else input.value = value ?? '';
        } else {
          writeInputValue(input, value);
        }
      }
      if (field === 'discountRate') input.placeholder = `project ${formatPercent(state.config.discount)}`;
      if (field === 'dutyRate') input.placeholder = `project ${formatPercent(state.freight.dutyRate)}`;
      if (field === 'insuranceRate') input.placeholder = `project ${formatPercent(state.freight.insuranceRate)}`;
      if (field === 'shippingAmount') {
        const mode = makerFieldValue(name, 'shippingMode');
        input.disabled = mode === 'sheet';
        input.placeholder = mode === 'sheet' ? 'sheet' : '0';
      }
    });
    document.querySelectorAll('[data-maker-show]').forEach((node) => {
      const name = (node.dataset.makerBind || node.querySelector?.('[data-maker-bind]')?.dataset.makerBind);
      if (!name) {
        const holder = node.closest('[data-maker]');
        if (!holder) return;
        node.hidden = (makerFieldValue(holder.dataset.maker, 'shippingMode') === 'percent') !== (node.dataset.makerShow === 'percent');
        return;
      }
      node.hidden = (makerFieldValue(name, 'shippingMode') === 'percent') !== (node.dataset.makerShow === 'percent');
    });
    document.querySelectorAll('[data-maker-rate]').forEach((node) => { node.textContent = makerRateText(node.dataset.makerRate); });
    document.querySelectorAll('[data-maker-count]').forEach((node) => {
      const count = makerKits(node.dataset.makerCount).length;
      node.textContent = `${count} kit${count === 1 ? '' : 's'}`;
    });
    document.querySelectorAll('[data-maker-route]').forEach((node) => {
      const sheet = makerKits(node.dataset.makerRoute)[0];
      if (sheet) node.textContent = kitRouteSummary(kitSettings(sheet));
    });
    document.querySelectorAll('[data-kit-currency]').forEach((node) => {
      const kit = kitSettings(node.dataset.kitCurrency);
      node.textContent = kit.shippingMode === 'percent' ? '%' : kit.shippingCurrency;
    });
  }

  function renameMaker(oldName, requested) {
    const next = String(requested || '').trim().slice(0, 80);
    if (!next || next === oldName) {
      syncMakerInputs();
      return;
    }
    makerKits(oldName).forEach((sheet) => { kitSettings(sheet).manufacturer = next; });
    if (state.makers[oldName]) {
      if (!state.makers[next]) state.makers[next] = state.makers[oldName];
      delete state.makers[oldName];
    }
    saveSettingsSoon();
    renderMakers();
    rebuildConsolidation();
  }

  function handleMakerInput(input, eventType) {
    const name = input.dataset.makerBind;
    const field = input.dataset.makerField;
    const discrete = input.tagName === 'SELECT';
    // Names are applied when the field is left (a rename rebuilds the cards); everything else follows each keystroke.
    if (field === 'name') {
      if (eventType === 'change') renameMaker(name, input.value);
      return;
    }
    if (eventType === 'change' && !discrete) return;
    if (field === 'location') {
      const maker = makerSettings(name);
      maker.location = String(input.value || '').slice(0, 80);
      saveSettingsSoon();
      return;
    }
    if (MAKER_KIT_FIELDS.has(field)) {
      const parsed = readKitField(field, input);
      if (!parsed) return;
      makerKits(name).forEach((sheet) => { kitSettings(sheet)[field] = parsed.value; });
    } else {
      const maker = makerSettings(name);
      if (field === 'shipmentMode') {
        maker.shipmentMode = input.value === 'own' ? 'own' : 'project';
      } else if (field === 'shipmentCurrency' || field === 'clearanceCurrency') {
        maker[field] = CURRENCIES.includes(input.value) ? input.value : MAKER_DEFAULTS[field];
      } else if (field === 'shipmentAmount' || field === 'clearanceAmount') {
        const value = readNumberInput(input);
        if (value === undefined) return;
        maker[field] = Math.max(0, value);
      } else if (field === 'insuranceRate') {
        if (input.value.trim() === '') {
          maker.insuranceRate = null;
        } else {
          const value = readNumberInput(input);
          if (value === undefined) return;
          maker.insuranceRate = clamp(value, 0, 1);
        }
      } else {
        return;
      }
    }
    saveSettingsSoon();
    if (discrete) {
      renderMakers();
      rebuildConsolidation();
      if (['shippingCurrency', 'defaultCurrency', 'shipmentCurrency', 'clearanceCurrency'].includes(field)) refreshRates(false);
      if (field === 'originCountry') renderMaps();
      updateKitRouteSummaries();
    } else {
      scheduleRecalc('consolidation');
    }
  }

  /* ---------- Factory-to-customer flow: one lane per manufacturer, one shared warehouse ---------- */

  // The control inside a manufacturer input's label, for use in a table cell.
  function makerCell(name, field, spec) {
    return makerControl(name, field, spec).lastElementChild;
  }

  function flowRow(label, ...cells) {
    const row = document.createElement('tr');
    const th = document.createElement('th');
    th.scope = 'row';
    if (typeof label === 'string') th.textContent = label;
    else th.append(label);
    const td = document.createElement('td');
    td.append(...cells);
    row.append(th, td);
    return row;
  }

  function amountWithCurrency(amount, currency) {
    const wrap = document.createElement('span');
    wrap.className = 'amount-currency compact';
    wrap.append(amount, currency);
    return wrap;
  }

  function boundControl(path, spec = {}) {
    let control;
    if (spec.kind === 'currency') {
      control = document.createElement('select');
      populateCurrencySelect(control, getPath(path));
    } else if (spec.kind === 'select') {
      control = document.createElement('select');
      spec.options.forEach(([value, text]) => control.append(new Option(text, value)));
    } else {
      control = document.createElement('input');
      control.type = 'number';
      control.min = '0';
      control.step = spec.step || '1';
      if (spec.max !== undefined) control.max = String(spec.max);
      if (spec.scale) control.dataset.scale = String(spec.scale);
    }
    control.dataset.bind = path;
    control.setAttribute('aria-label', spec.label || path);
    return control;
  }

  // A small labelled field: the caption sits above the control, so two fields fit side by side in a narrow box.
  function miniField(label, controls, { wide = false, extra = false, info = false } = {}) {
    const field = document.createElement('div');
    field.className = `mini-field${wide ? ' wide' : ''}${extra ? ' maker-extra' : ''}`;
    const caption = document.createElement('span');
    caption.className = 'mini-label';
    caption.textContent = label;
    const body = document.createElement('div');
    body.className = 'mini-controls';
    body.append(...[].concat(controls));
    field.append(caption, body);
    if (info) field.dataset.customsInfo = '';
    return field;
  }

  function miniLine(label) {
    const line = document.createElement('div');
    line.className = 'mini-total';
    const text = document.createElement('span');
    text.textContent = label;
    const value = document.createElement('strong');
    line.append(text, value);
    return { line, value };
  }

  function buildKitFreightRow(sheetName) {
    const row = document.createElement('div');
    row.className = 'mini-kit';
    const name = document.createElement('span');
    name.className = 'mini-label';
    name.textContent = sheetName;
    name.translate = false;
    const wrap = document.createElement('div');
    wrap.className = 'kit-inline';
    const mode = document.createElement('select');
    mode.dataset.kitBind = sheetName;
    mode.dataset.kitField = 'shippingMode';
    mode.setAttribute('aria-label', `${sheetName} freight mode`);
    Object.entries(KIT_SHIPPING_MODES).forEach(([value, label]) => mode.append(new Option(label, value)));
    const amount = document.createElement('input');
    amount.type = 'number';
    amount.min = '0';
    amount.step = '0.01';
    amount.dataset.kitBind = sheetName;
    amount.dataset.kitField = 'shippingAmount';
    amount.setAttribute('aria-label', `${sheetName} freight amount`);
    const currency = document.createElement('span');
    currency.className = 'kit-currency';
    currency.dataset.kitCurrency = sheetName;
    wrap.append(mode, amount, currency);
    row.append(name, wrap);
    return row;
  }

  // The editable "Shipment & tolls" box of one manufacturer's route. Collapsed it shows the two numbers that matter
  // most and the total; opened it shows everything, including the freight per kit.
  function buildMakerBox(name) {
    const own = makerHasOwnShipment(name);
    const box = document.createElement('section');
    box.className = 'flow-box maker-box';
    box.dataset.maker = name;
    box.dataset.makerBox = name;
    box.setAttribute('aria-label', 'Shipment and tolls');
    const heading = document.createElement('h4');
    const title = document.createElement('span');
    title.textContent = 'Shipment & tolls';
    heading.append(title, buildBoxToggle(`ship:${name}`, 'Show or hide the freight per kit'));
    const who = document.createElement('p');
    who.className = 'maker-box-name';
    who.translate = false;
    who.textContent = name;

    const grid = document.createElement('div');
    grid.className = 'mini-grid';
    const shipment = own
      ? amountWithCurrency(makerCell(name, 'shipmentAmount', { kind: 'number', caption: 'Shipment', step: '1' }), makerCell(name, 'shipmentCurrency', { kind: 'currency', caption: 'Shipment currency' }))
      : amountWithCurrency(
        boundControl('freight.consolidatedShipment', { label: 'Consolidated shipment amount' }),
        boundControl('freight.consolidatedCurrency', { kind: 'currency', label: 'Consolidated shipment currency' }),
      );
    const insurance = own
      ? makerCell(name, 'insuranceRate', { kind: 'number', caption: 'Insurance', suffix: '%', scale: 100, max: 100, step: '0.1' })
      : (() => {
        const wrap = document.createElement('span');
        wrap.className = 'input-suffix';
        const unit = document.createElement('span');
        unit.textContent = '%';
        wrap.append(boundControl('freight.insuranceRate', { label: 'Insurance', scale: 100, max: 100, step: '0.1' }), unit);
        return wrap;
      })();
    const clearance = own
      ? amountWithCurrency(makerCell(name, 'clearanceAmount', { kind: 'number', caption: 'Clearance and broker fees', step: '1' }), makerCell(name, 'clearanceCurrency', { kind: 'currency', caption: 'Clearance fee currency' }))
      : amountWithCurrency(boundControl('freight.clearanceFee', { label: 'Clearance and broker fees' }), boundControl('freight.clearanceCurrency', { kind: 'currency', label: 'Clearance fee currency' }));
    const kits = document.createElement('div');
    kits.className = 'mini-field wide maker-extra maker-kit-rows';
    const kitsCaption = document.createElement('span');
    kitsCaption.className = 'mini-label';
    kitsCaption.textContent = 'Freight per kit';
    kits.append(kitsCaption, ...makerKits(name).map(buildKitFreightRow));
    const vat = miniLine('Import VAT (deductible)');
    vat.line.classList.add('maker-extra');
    vat.line.dataset.customsInfo = '';
    const total = miniLine('Freight & import total');
    grid.append(
      miniField('Shipment', shipment),
      miniField('Customs duty', makerCell(name, 'dutyRate', { kind: 'number', caption: 'Customs duty', suffix: '%', scale: 100, max: 100, step: '0.1' }), { info: true }),
      miniField('Insurance', insurance, { extra: true }),
      miniField('Consolidated shipment', makerCell(name, 'shipmentMode', { kind: 'select', caption: 'Consolidated shipment', options: MAKER_SHIPMENT_OPTIONS }), { wide: true, extra: true }),
      miniField('Split by', boundControl('freight.allocation', { kind: 'select', label: 'Split consolidated shipment by', options: [['value', 'Value'], ['quantity', 'Quantity'], ['lines', 'Lines']] }), { extra: true }),
      miniField('Clearance & broker', clearance, { wide: true, extra: true, info: true }),
      kits,
      vat.line,
      total.line,
    );
    box.append(heading, who, grid);
    return { box, vatCell: vat.value, totalCell: total.value };
  }

  function buildMakerLane(name) {
    const lane = document.createElement('div');
    lane.className = 'flow-lane';
    lane.dataset.maker = name;

    const factory = document.createElement('article');
    factory.className = 'flow-node factory compact';
    const kind = document.createElement('div');
    kind.className = 'flow-kind';
    kind.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 21V10l5 3V10l5 3V6l8 4v11z"/></svg>';
    kind.append('Manufacturer');
    const nameField = makerCell(name, 'name', { kind: 'text', caption: 'Manufacturer name' });
    nameField.classList.add('flow-title-input');
    const place = makerCell(name, 'location', { kind: 'text', caption: 'Manufacturer location' });
    const country = makerCell(name, 'originCountry', { kind: 'country', country: true, caption: 'Country' });
    const freightAmount = makerCell(name, 'shippingAmount', { kind: 'number', caption: 'Freight amount', step: '0.01' });
    const freightUnit = document.createElement('small');
    freightUnit.className = 'maker-unit';
    freightUnit.dataset.makerShow = 'percent';
    freightUnit.textContent = '% of the item value';
    const rate = document.createElement('small');
    rate.className = 'maker-rate';
    rate.dataset.makerRate = name;
    const stats = document.createElement('dl');
    stats.className = 'stat-list';
    const grid = document.createElement('div');
    grid.className = 'mini-grid';
    const countryField = miniField('Country', country, { info: true });
    grid.append(
      miniField('Location', place),
      countryField,
      miniField('Discount', makerCell(name, 'discountRate', { kind: 'number', caption: 'Discount', suffix: '%', scale: 100, max: 100, step: '0.1' })),
      miniField('Rate buffer', makerCell(name, 'fxMargin', { kind: 'number', caption: 'Rate buffer', step: '0.01' })),
      miniField('Freight add-on', makerCell(name, 'shippingMode', { kind: 'select', caption: 'Freight add-on', options: KIT_MODE_OPTIONS() })),
      miniField('Amount', [freightAmount, freightUnit]),
    );
    factory.append(kind, nameField, grid, rate, stats);

    const leg = document.createElement('div');
    leg.className = 'flow-leg';
    leg.dataset.makerLeg = name;
    const arrow = document.createElement('div');
    arrow.className = 'flow-arrow';
    const arrowText = document.createElement('span');
    arrow.append(arrowText);
    leg.append(arrow);

    const border = document.createElement('article');
    border.className = 'flow-node border compact';
    const borderKind = document.createElement('div');
    borderKind.className = 'flow-kind';
    borderKind.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h2v18H5zM8 4h11l-2 4 2 4H8z"/></svg>';
    const destName = document.createElement('span');
    borderKind.append(destName);
    const borderTitle = document.createElement('strong');
    borderTitle.className = 'flow-title';
    borderTitle.textContent = 'Customs & import';
    const landed = document.createElement('span');
    landed.className = 'flow-sub';
    const treatment = document.createElement('span');
    treatment.className = 'flow-sub maker-route';
    treatment.dataset.makerRoute = name;
    const borderStats = document.createElement('dl');
    borderStats.className = 'stat-list';
    border.append(borderKind, borderTitle, landed, treatment, borderStats);

    lane.append(factory, leg, border);
    const { box, vatCell, totalCell } = buildMakerBox(name);
    return { lane, refs: { box, leg, arrowText, destName, landed, factoryStats: stats, borderStats, vatCell, totalCell } };
  }

  // "More / Less" button that folds a box's rarely used fields away.
  function buildBoxToggle(key, label) {
    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'box-toggle';
    toggle.dataset.mapBoxToggle = key;
    toggle.setAttribute('aria-expanded', 'false');
    toggle.setAttribute('aria-label', label);
    toggle.title = label;
    toggle.textContent = 'More';
    return toggle;
  }

  function setBoxToggle(box, open) {
    const toggle = box.querySelector('[data-map-box-toggle]');
    if (!toggle) return;
    toggle.setAttribute('aria-expanded', String(Boolean(open)));
    toggle.textContent = open ? 'Less' : 'More';
  }

  function allFlowBoxes() {
    return [...[...state.laneRefs.values()].map((refs) => refs.box), dom.flowBoxWarehouse, dom.flowBoxCustomer];
  }

  // The editable tables sit in the route row in Route view and on the route itself in Map view.
  function placeFlowBoxes() {
    const onMap = state.flowMode === 'map';
    const target = (leg) => (onMap ? dom.mapBoxes : leg);
    state.laneRefs.forEach((refs) => target(refs.leg).append(refs.box));
    target(dom.flowLegWarehouse).append(dom.flowBoxWarehouse);
    target(dom.flowLegCustomer).append(dom.flowBoxCustomer);
    if (!onMap) resetFlowBoxes();
  }

  function resetFlowBoxes() {
    allFlowBoxes().forEach((box) => {
      box.style.left = '';
      box.style.top = '';
      // The manufacturers' tables keep their own open or closed state in both views.
      if (state.flowMode !== 'map' && !box.dataset.makerBox) box.classList.remove('is-collapsed');
    });
  }

  function renderFlowLanes() {
    const names = makerNames();
    const signature = names.map((name) => `${name}|${makerKits(name).join(',')}|${makerSettings(name).shipmentMode}`).join('\u0001');
    if (dom.flowLanes.dataset.signature !== signature) {
      dom.flowLanes.dataset.signature = signature;
      state.laneRefs = new Map();
      const lanes = names.map((name) => {
        const { lane, refs } = buildMakerLane(name);
        state.laneRefs.set(name, refs);
        return lane;
      });
      dom.flowLanes.replaceChildren(...lanes);
      placeFlowBoxes();
    }
    const destination = destinationConfig();
    const sum = (list, field) => list.reduce((total, item) => total + (item[field] ?? 0), 0);
    state.laneRefs.forEach((refs, name) => {
      const items = state.consolidated.filter((item) => (item.manufacturer || '') === name);
      const units = items.reduce((total, item) => total + Math.max(0, item.maxQuantity || 0), 0);
      refs.arrowText.textContent = `Ship to ${destination.name}`;
      refs.destName.textContent = destination.name;
      refs.landed.textContent = `Landed in ${destination.name}`;
      setStats(refs.factoryStats, [
        ['List value', formatWhole(sum(items, 'convertedTotal'))],
        ['After discount', formatWhole(sum(items, 'discountedTotal'))],
        ['Parts · units', `${items.length} · ${formatNumber(units, 0)}`],
      ]);
      setStats(refs.borderStats, [
        ['Landed value', formatWhole(sum(items, 'landedCost'))],
        ['Freight & import', formatWhole(sum(items, 'freightAndImport'))],
        ['Import VAT', formatWhole(sum(items, 'importVat'))],
      ]);
      refs.vatCell.textContent = formatWhole(sum(items, 'importVat'));
      refs.totalCell.textContent = formatWhole(sum(items, 'freightAndImport'));
      // The per-kit rows and the rarely used fields start folded away; the toggle opens them in either view.
      const open = Boolean(state.map.open[`ship:${name}`]);
      refs.box.classList.toggle('is-collapsed', !open);
      setBoxToggle(refs.box, open);
    });
    syncMakerInputs();
  }

  /* ---------- Settings persistence (local + workbook + secure project) ---------- */

  function collectSettings() {
    const warehouse = { ...state.warehouse, rates: { ...state.warehouse.rates } };
    delete warehouse.shareMessage;
    const mappings = {};
    Object.entries(state.mappings).forEach(([sheetName, mapping]) => {
      if (isGeneratedSheet(state.workbook?.Sheets[sheetName])) return;
      mappings[sheetName] = { ...mapping };
    });
    return {
      format: 'partslist-settings',
      version: 3,
      language: state.language,
      config: {
        discount: state.config.discount,
        multiplier: state.config.multiplier,
        shippingMargin: state.config.shippingMargin,
        vatRate: state.config.vatRate,
        sourceCurrency: state.config.sourceCurrency,
        outputCurrency: state.config.outputCurrency,
      },
      freight: { ...state.freight },
      kits: JSON.parse(JSON.stringify(state.kits)),
      makers: JSON.parse(JSON.stringify(state.makers)),
      groups: state.groups.map((group) => ({ ...group })),
      items: JSON.parse(JSON.stringify(state.items)),
      sales: { ...state.sales },
      scenario: { ...state.scenario },
      supplier: { ...state.supplier },
      customer: { ...state.customer },
      vat: { byCountry: { ...state.vat.byCountry } },
      flowOverrides: { ...state.flowOverrides },
      wireNodePositions: Core.sanitizeWirePositions(state.wireNodePositions),
      wireRelations: state.wireRelations.filter(isValidWireRelation).map((edge) => ({ ...edge })),
      mappings,
      variables: state.variables.map((variable) => ({ name: variable.name, expression: variable.expression })),
      warehouse,
      disabledKits: state.workbook ? state.workbook.SheetNames.filter((name) => isSourceKit(name) && !state.mappings[name].enabled) : [],
    };
  }

  function finiteOr(value, fallback, min = -Infinity, max = Infinity) {
    const number = typeof value === 'number' ? value : toNumber(value);
    return Number.isFinite(number) ? clamp(number, min, max) : fallback;
  }

  function nullableFraction(value) {
    const number = nullableNumber(value, 0);
    return number === null ? null : Math.min(1, number);
  }

  function nullableNumber(value, min = 0) {
    if (value === null || value === undefined || value === '') return null;
    const number = typeof value === 'number' ? value : toNumber(value);
    return Number.isFinite(number) ? Math.max(min, number) : null;
  }

  function applySettings(settings) {
    if (!settings || typeof settings !== 'object' || settings.format !== 'partslist-settings') return false;
    if (settings.language === 'en' || settings.language === 'sv') state.language = settings.language;
    const config = settings.config || {};
    state.config.discount = finiteOr(config.discount, state.config.discount, 0, 1);
    state.config.multiplier = finiteOr(config.multiplier, state.config.multiplier, 0);
    state.config.shippingMargin = finiteOr(config.shippingMargin, state.config.shippingMargin, 0, 0.999);
    state.config.vatRate = finiteOr(config.vatRate, state.config.vatRate, 0, 1);
    if (CURRENCIES.includes(config.sourceCurrency)) state.config.sourceCurrency = config.sourceCurrency;
    if (CURRENCIES.includes(config.outputCurrency)) state.config.outputCurrency = config.outputCurrency;

    const freight = settings.freight || {};
    state.freight = {
      consolidatedShipment: finiteOr(freight.consolidatedShipment, 0, 0),
      consolidatedCurrency: CURRENCIES.includes(freight.consolidatedCurrency) ? freight.consolidatedCurrency : FREIGHT_DEFAULTS.consolidatedCurrency,
      allocation: ['value', 'quantity', 'lines'].includes(freight.allocation) ? freight.allocation : 'value',
      dutyRate: finiteOr(freight.dutyRate, 0, 0, 1),
      insuranceRate: finiteOr(freight.insuranceRate, 0, 0, 1),
      clearanceFee: finiteOr(freight.clearanceFee, 0, 0),
      clearanceCurrency: CURRENCIES.includes(freight.clearanceCurrency) ? freight.clearanceCurrency : FREIGHT_DEFAULTS.clearanceCurrency,
    };

    state.kits = {};
    Object.entries(settings.kits || {}).forEach(([name, kit]) => {
      if (!kit || typeof kit !== 'object') return;
      state.kits[name] = {
        shippingMode: KIT_SHIPPING_MODES[kit.shippingMode] ? kit.shippingMode : 'sheet',
        shippingAmount: finiteOr(kit.shippingAmount, KIT_DEFAULTS.shippingAmount, 0),
        shippingCurrency: CURRENCIES.includes(kit.shippingCurrency) ? kit.shippingCurrency : KIT_DEFAULTS.shippingCurrency,
        manufacturer: typeof kit.manufacturer === 'string' ? kit.manufacturer.trim().slice(0, 80) || name : name,
        originCountry: Core.normalizeCountryCode(kit.originCountry, state.supplier.country),
        defaultCurrency: CURRENCIES.includes(kit.defaultCurrency) ? kit.defaultCurrency : null,
        discountRate: nullableFraction(kit.discountRate),
        dutyRate: nullableFraction(kit.dutyRate),
        fxMargin: nullableNumber(kit.fxMargin),
      };
    });
    state.makers = {};
    Object.entries(settings.makers || {}).forEach(([name, maker]) => {
      if (!maker || typeof maker !== 'object') return;
      state.makers[name] = sanitizeMaker(maker);
    });

    state.groups = (Array.isArray(settings.groups) ? settings.groups : []).slice(0, MAX_GROUPS)
      .filter((group) => group && typeof group.id === 'string' && typeof group.name === 'string')
      .map((group, index) => ({
        id: group.id.slice(0, 40),
        name: group.name.trim().slice(0, 40) || `Group ${index + 1}`,
        color: GROUP_COLORS.includes(group.color) ? group.color : GROUP_COLORS[index],
        salesMode: group.salesMode === 'units' ? 'units' : 'turns',
        salesValue: nullableNumber(group.salesValue),
        multiplier: nullableNumber(group.multiplier),
      }));
    const groupIds = new Set(state.groups.map((group) => group.id));

    state.items = {};
    Object.entries(settings.items || {}).forEach(([key, item]) => {
      if (!item || typeof item !== 'object') return;
      const clean = {
        group: groupIds.has(item.group) ? item.group : null,
        salesPerYear: nullableNumber(item.salesPerYear),
        quantity: nullableNumber(item.quantity),
        multiplier: nullableNumber(item.multiplier),
        excluded: item.excluded === true,
      };
      if (clean.group || clean.salesPerYear !== null || clean.quantity !== null || clean.multiplier !== null || clean.excluded) state.items[key] = clean;
    });

    const sales = settings.sales || {};
    state.sales = {
      defaultMode: sales.defaultMode === 'units' ? 'units' : 'turns',
      defaultValue: finiteOr(sales.defaultValue, SALES_DEFAULTS.defaultValue, 0),
      horizonMonths: Math.round(finiteOr(sales.horizonMonths, SALES_DEFAULTS.horizonMonths, 6, 120)),
      customerPaysOutbound: sales.customerPaysOutbound === true,
    };
    const scenario = settings.scenario || {};
    state.scenario = {
      salesDelta: finiteOr(scenario.salesDelta, 0, -100, 500),
      purchaseDelta: finiteOr(scenario.purchaseDelta, 0, -100, 500),
      inboundDelta: finiteOr(scenario.inboundDelta, 0, -100, 500),
      warehouseDelta: finiteOr(scenario.warehouseDelta, 0, -100, 500),
    };
    state.customer = { country: DESTINATIONS[settings.customer?.country] ? settings.customer.country : '578' };
    state.vat = { unlocked: false, byCountry: { ...VAT_DEFAULTS } };
    Object.keys(DESTINATIONS).forEach((id) => {
      const saved = settings.vat?.byCountry?.[id];
      if (VAT_TREATMENT_KEYS.includes(saved)) state.vat.byCountry[id] = saved;
    });
    const supplier = settings.supplier || {};
    const country = Core.normalizeCountryCode(supplier.country, SUPPLIER_DEFAULTS.country);
    state.supplier = {
      name: typeof supplier.name === 'string' ? supplier.name.slice(0, 80) : SUPPLIER_DEFAULTS.name,
      location: typeof supplier.location === 'string' ? supplier.location.slice(0, 80) : SUPPLIER_DEFAULTS.location,
      country,
      countryName: typeof supplier.countryName === 'string' ? supplier.countryName.slice(0, 80) : ORIGIN_PRESETS[country]?.name || '',
      lat: finiteOr(supplier.lat, ORIGIN_PRESETS[country]?.lat ?? SUPPLIER_DEFAULTS.lat, -90, 90),
      lon: finiteOr(supplier.lon, ORIGIN_PRESETS[country]?.lon ?? SUPPLIER_DEFAULTS.lon, -180, 180),
    };
    FLOW_OVERRIDE_KEYS.forEach((key) => { state.flowOverrides[key] = nullableNumber(settings.flowOverrides?.[key]); });
    state.wireNodePositions = Core.sanitizeWirePositions(settings.wireNodePositions);

    const warehouse = settings.warehouse || {};
    Object.keys(WAREHOUSE_DEFAULTS).forEach((key) => {
      const fallback = WAREHOUSE_DEFAULTS[key];
      if (!(key in warehouse)) return;
      if (typeof fallback === 'number') state.warehouse[key] = finiteOr(warehouse[key], state.warehouse[key], 0);
      else if (typeof fallback === 'boolean') state.warehouse[key] = warehouse[key] === true;
      else if (typeof warehouse[key] === 'string') state.warehouse[key] = warehouse[key].slice(0, 80);
    });
    state.warehouse.palletType = state.warehouse.palletType === 'sea' ? 'sea' : 'eu';
    state.warehouse.country = DESTINATIONS[state.warehouse.country] ? state.warehouse.country : WAREHOUSE_DEFAULTS.country;
    state.warehouse.palletHeight = state.warehouse.palletHeight === '220' ? '220' : '120';
    state.warehouse.binsPerShelf = clamp(Math.round(state.warehouse.binsPerShelf) || 1, 1, 12);
    state.warehouse.shelvesPerRack = clamp(Math.round(state.warehouse.shelvesPerRack) || 1, 1, 8);
    state.warehouse.unitsPerShelf = state.warehouse.binsPerShelf * state.warehouse.unitsPerBin;
    Object.keys(WAREHOUSE_RATE_DEFAULTS).forEach((key) => {
      if (settings.warehouse?.rates && key in settings.warehouse.rates) {
        state.warehouse.rates[key] = finiteOr(settings.warehouse.rates[key], state.warehouse.rates[key], 0);
      }
    });

    state.pendingDisabledKits = Array.isArray(settings.disabledKits) ? settings.disabledKits.filter((name) => typeof name === 'string') : null;
    state.pendingMappings = settings.mappings && typeof settings.mappings === 'object' ? settings.mappings : null;
    state.pendingVariables = Array.isArray(settings.variables) ? settings.variables : null;
    state.pendingWireRelations = Array.isArray(settings.wireRelations) ? settings.wireRelations.filter(isValidWireRelation) : null;
    return true;
  }

  function applyPendingMappings() {
    if (!state.pendingMappings || !state.workbook) return;
    const numericFields = ['headerRow', 'dataStartRow', ...Object.keys(FIELD_LABELS)];
    Object.entries(state.pendingMappings).forEach(([sheetName, saved]) => {
      const mapping = state.mappings[sheetName];
      const matrix = state.matrices[sheetName];
      if (!mapping || !matrix || !saved || typeof saved !== 'object') return;
      if (typeof saved.enabled === 'boolean') mapping.enabled = saved.enabled;
      numericFields.forEach((field) => {
        if (saved[field] === null && field in FIELD_LABELS) mapping[field] = null;
        else if (Number.isInteger(saved[field]) && saved[field] >= 0) mapping[field] = saved[field];
      });
      mapping.headerRow = clamp(mapping.headerRow, 0, Math.max(0, matrix.length - 1));
      mapping.dataStartRow = clamp(mapping.dataStartRow, mapping.headerRow + 1, Math.max(mapping.headerRow + 1, matrix.length - 1));
    });
    state.pendingMappings = null;
  }

  function restorePendingWorkbookState() {
    if (state.pendingVariables) restoreVariables(state.pendingVariables);
    if (state.pendingWireRelations) state.wireRelations = state.pendingWireRelations.map((edge) => ({ ...edge }));
    state.pendingVariables = null;
    state.pendingWireRelations = null;
  }

  function applyPendingDisabledKits() {
    if (!state.pendingDisabledKits || !state.workbook) return;
    state.pendingDisabledKits.forEach((name) => { if (state.mappings[name]) state.mappings[name].enabled = false; });
    state.pendingDisabledKits = null;
  }

  function saveSettingsSoon() {
    // Business data is intentionally not written to browser storage. Track the
    // unsaved state so users know when to export a workbook or secure project.
    markProjectDirty();
  }

  function markProjectDirty() {
    if (!state.workbook) return;
    state.projectDirty = true;
    state.projectSaveLabel = '';
    renderSaveState();
  }

  /* ---------- Undo & redo ---------- */

  // A snapshot is only the small, reversible slice of the project: the settings the user edits by hand,
  // the per-item overrides, the sales groups and the named variables.
  function captureHistoryState() {
    return {
      config: { ...state.config },
      freight: { ...state.freight },
      sales: { ...state.sales },
      scenario: { ...state.scenario },
      customer: { ...state.customer },
      supplier: { ...state.supplier },
      vat: { unlocked: state.vat.unlocked, byCountry: { ...state.vat.byCountry } },
      kits: JSON.parse(JSON.stringify(state.kits)),
      makers: JSON.parse(JSON.stringify(state.makers)),
      groups: JSON.parse(JSON.stringify(state.groups)),
      items: JSON.parse(JSON.stringify(state.items)),
      flowOverrides: { ...state.flowOverrides },
      variables: state.variables.map((variable) => ({ ...variable })),
      disabledKits: Object.entries(state.mappings).filter(([, mapping]) => !mapping.enabled).map(([name]) => name),
    };
  }

  function restoreHistoryState(snapshot) {
    if (!snapshot) return;
    state.config = { ...state.config, ...snapshot.config };
    state.freight = { ...snapshot.freight };
    state.sales = { ...snapshot.sales };
    state.scenario = { ...snapshot.scenario };
    state.customer = { ...snapshot.customer };
    state.supplier = { ...snapshot.supplier };
    state.vat = { unlocked: snapshot.vat.unlocked, byCountry: { ...snapshot.vat.byCountry } };
    state.kits = JSON.parse(JSON.stringify(snapshot.kits));
    state.makers = JSON.parse(JSON.stringify(snapshot.makers || {}));
    state.groups = JSON.parse(JSON.stringify(snapshot.groups));
    state.items = JSON.parse(JSON.stringify(snapshot.items));
    state.flowOverrides = { ...snapshot.flowOverrides };
    state.variables = (snapshot.variables || []).map((variable) => ({ ...variable }));
    Object.values(state.mappings).forEach((mapping) => { mapping.enabled = true; });
    (snapshot.disabledKits || []).forEach((name) => {
      if (state.mappings[name]) state.mappings[name].enabled = false;
    });
  }

  // Records a reversible edit. Call this *before* changing state so the snapshot holds the "before" value;
  // the action itself runs immediately, and is kept so redo can replay exactly the same edit.
  function pushHistory(label, apply, options = {}) {
    if (state.suppressHistory || !state.workbook) {
      apply();
      return;
    }
    const before = captureHistoryState();
    apply();
    state.undoStack.push({ label, before, apply, restore: options.restore || null });
    if (state.undoStack.length > UNDO_LIMIT) state.undoStack.shift();
    state.redoStack = [];
    state.undoLabel = label;
    renderUndoState();
  }

  function undo() {
    const entry = state.undoStack.pop();
    if (!entry) return;
    // The snapshot holds the state from before the edit, so restoring it and replaying the action on redo
    // moves the project back and forth without either side re-deriving the other's values.
    state.suppressHistory = true;
    restoreHistoryState(entry.before);
    // Worksheet cells live outside the settings snapshot, so their restore step runs separately.
    entry.restore?.();
    state.suppressHistory = false;
    state.redoStack.push(entry);
    state.undoLabel = entry.label;
    // Kits switched on or off change which parts are consolidated, so the list is rebuilt, not just redrawn.
    renderMappings();
    rebuildConsolidation();
    refreshRates(false);
    renderGroups();
    renderVariables();
    renderUndoState();
    showToast(`Undone: ${entry.label}.`);
  }

  function redo() {
    const entry = state.redoStack.pop();
    if (!entry) return;
    state.suppressHistory = true;
    try {
      // A redo follows an undo, so worksheet cells the edit overwrote still hold their restored values.
      // apply() writes the edit on top of them, which is exactly what redoing the edit should mean.
      entry.apply();
    } finally {
      state.suppressHistory = false;
    }
    state.undoStack.push(entry);
    state.undoLabel = entry.label;
    // Kits switched on or off change which parts are consolidated, so the list is rebuilt, not just redrawn.
    renderMappings();
    rebuildConsolidation();
    refreshRates(false);
    renderGroups();
    renderVariables();
    renderUndoState();
    showToast(`Redone: ${entry.label}.`);
  }

  function renderUndoState() {
    const canUndo = state.undoStack.length > 0;
    const canRedo = state.redoStack.length > 0;
    dom.undoButton.disabled = !canUndo;
    const label = canUndo ? state.undoLabel : 'Undo';
    const next = canRedo ? 'Redo' : label;
    dom.undoButton.title = canUndo
      ? `${canRedo ? 'Redo' : 'Undo'} ${label} (${canRedo ? 'Ctrl+Shift+Z' : 'Ctrl+Z'})`
      : 'Nothing to undo (Ctrl+Z)';
    dom.undoButton.setAttribute('aria-label', canUndo ? `${next} ${label}` : 'Nothing to undo');
    dom.undoButton.querySelector('.undo-label').textContent = next;
  }

  function markProjectSaved(label) {
    state.projectDirty = false;
    state.projectSaveLabel = label || 'Saved';
    renderSaveState();
  }

  function renderSaveState() {
    renderUndoState();
    if (!state.workbook) {
      dom.saveState.textContent = 'No workbook loaded';
      dom.saveState.className = 'save-state';
      dom.guideToggle.disabled = true;
      return;
    }
    dom.guideToggle.disabled = false;
    dom.saveState.textContent = state.projectDirty ? 'Unsaved changes' : state.projectSaveLabel || 'Workbook loaded';
    dom.saveState.className = `save-state ${state.projectDirty ? 'dirty' : 'saved'}`;
  }

  function renderSidebarState() {
    const collapsed = state.sidebarCollapsed;
    const label = collapsed ? 'Show settings' : 'Hide settings';
    dom.workspace.classList.toggle('sidebar-collapsed', collapsed);
    dom.sidebarToggle.setAttribute('aria-expanded', String(!collapsed));
    dom.sidebarToggle.setAttribute('aria-label', label);
    dom.sidebarToggle.title = label;
    dom.sidebarCollapse.setAttribute('aria-expanded', String(!collapsed));
    dom.sidebarExpand.setAttribute('aria-expanded', String(!collapsed));
    dom.sidebarBody.toggleAttribute('inert', collapsed);
    dom.sidebarBody.setAttribute('aria-hidden', String(collapsed));
    // Move focus to the control that is still visible so keyboard users are not left on a hidden button.
    if (document.activeElement === dom.sidebarCollapse && collapsed) dom.sidebarExpand.focus();
    else if (document.activeElement === dom.sidebarExpand && !collapsed) dom.sidebarCollapse.focus();
  }

  // Only harmless, non-business interface choices are remembered in the browser.
  function loadUiPreferences() {
    try {
      const stored = JSON.parse(localStorage.getItem(UI_PREFERENCES_KEY) || 'null');
      if (stored?.language === 'en' || stored?.language === 'sv') state.language = stored.language;
      if (typeof stored?.sidebarCollapsed === 'boolean') state.sidebarCollapsed = stored.sidebarCollapsed;
      if (stored?.consolidatedSort && ['asc', 'desc'].includes(stored.consolidatedSort.direction) && typeof stored.consolidatedSort.key === 'string') {
        state.consolidatedSort = { key: stored.consolidatedSort.key, direction: stored.consolidatedSort.direction };
      }
      if (Array.isArray(stored?.hiddenColumns)) state.hiddenColumns = stored.hiddenColumns.filter((key) => typeof key === 'string');
    } catch { /* Preferences are optional. */ }
  }

  function saveUiPreferences() {
    try {
      // Only harmless, non-business interface choices are remembered: language, panel state and how the
      // parts table is arranged. No workbook or pricing data is written here.
      localStorage.setItem(UI_PREFERENCES_KEY, JSON.stringify({
        language: state.language,
        sidebarCollapsed: state.sidebarCollapsed,
        consolidatedSort: state.consolidatedSort,
        hiddenColumns: state.hiddenColumns,
      }));
    } catch { /* Preferences are optional. */ }
  }

  function purgeLegacyLocalBusinessData() {
    try {
      const keys = [];
      for (let index = 0; index < localStorage.length; index += 1) {
        const key = localStorage.key(index);
        if (key?.startsWith(LEGACY_SETTINGS_STORAGE_PREFIX) || key === WAREHOUSE_RATE_STORAGE_KEY) keys.push(key);
      }
      keys.forEach((key) => localStorage.removeItem(key));
    } catch (error) {
      console.warn('Legacy local business data could not be cleared.', error);
    }
  }

  function resetProjectSettings() {
    if (!state.workbook) return;
    if (!window.confirm(translateUi('Reset kit freight, freight & import, VAT, sales groups, item overrides and sales assumptions to their defaults? Warehouse quote rates are kept.'))) return;
    pushHistory('Settings reset to defaults', () => {
      state.config.vatRate = 0;
      state.vat = { unlocked: false, byCountry: { ...VAT_DEFAULTS } };
      state.freight = { ...FREIGHT_DEFAULTS };
      state.kits = {};
      state.makers = {};
      state.groups = [];
      state.items = {};
      state.sales = { ...SALES_DEFAULTS };
      state.scenario = { salesDelta: 0, purchaseDelta: 0, inboundDelta: 0, warehouseDelta: 0 };
      state.supplier = { ...SUPPLIER_DEFAULTS };
      state.customer = { country: '578' };
      FLOW_OVERRIDE_KEYS.forEach((key) => { state.flowOverrides[key] = null; });
      state.selectedKeys.clear();
      renderMappings();
      renderGroups();
      renderAllSelectionBars();
      rebuildConsolidation();
      refreshRates(false);
    });
    showToast('Kit, freight, group and sales settings were reset. Ctrl+Z restores them.');
  }

  function rotatePrintableAscii(value, shift) {
    return Array.from(String(value), (character) => {
      const code = character.charCodeAt(0);
      if (code < PRINTABLE_ASCII_START || code >= PRINTABLE_ASCII_START + PRINTABLE_ASCII_RANGE) return character;
      const offset = ((code - PRINTABLE_ASCII_START + shift) % PRINTABLE_ASCII_RANGE + PRINTABLE_ASCII_RANGE) % PRINTABLE_ASCII_RANGE;
      return String.fromCharCode(PRINTABLE_ASCII_START + offset);
    }).join('');
  }

  function createWarehouseRateEnvelope(rates = state.warehouse.rates) {
    const normalizedRates = {};
    Object.keys(WAREHOUSE_RATE_DEFAULTS).forEach((key) => {
      normalizedRates[key] = Math.max(0, toNumber(rates[key]) ?? 0);
    });
    const payload = {
      format: 'partslist-warehouse-rate-data',
      version: 1,
      currency: 'NOK',
      savedAt: new Date().toISOString(),
      rates: normalizedRates,
    };
    const shifted = rotatePrintableAscii(JSON.stringify(payload), CAESAR_SHIFT);
    return {
      format: 'partslist-warehouse-rates',
      version: 1,
      cipher: { name: 'CAESAR-PRINTABLE-ASCII', shift: CAESAR_SHIFT },
      payload: bytesToBase64(new TextEncoder().encode(shifted)),
    };
  }

  function decodeWarehouseRateEnvelope(envelope) {
    if (envelope?.format !== 'partslist-warehouse-rates' || envelope.version !== 1) {
      throw new Error('This is not a supported warehouse-rate JSON file.');
    }
    if (envelope.cipher?.name !== 'CAESAR-PRINTABLE-ASCII' || envelope.cipher?.shift !== CAESAR_SHIFT) {
      throw new Error('The rate file must use Caesar-14 encoding.');
    }
    const shifted = new TextDecoder().decode(base64ToBytes(envelope.payload));
    const payload = JSON.parse(rotatePrintableAscii(shifted, -CAESAR_SHIFT));
    if (payload?.format !== 'partslist-warehouse-rate-data' || payload.version !== 1 || payload.currency !== 'NOK') {
      throw new Error('The decoded warehouse-rate payload is invalid.');
    }
    const rates = {};
    Object.keys(WAREHOUSE_RATE_DEFAULTS).forEach((key) => {
      const value = toNumber(payload.rates?.[key]);
      if (!Number.isFinite(value) || value < 0) throw new Error(`The rate “${key}” is missing or invalid.`);
      rates[key] = value;
    });
    return rates;
  }

  function applyWarehouseRates(rates, statusText = 'Loaded for this session · export or save a secure project to keep') {
    state.warehouse.rates = { ...WAREHOUSE_RATE_DEFAULTS, ...rates };
    syncWarehouseRateInputs();
    if (dom.warehouseRatesStatus) dom.warehouseRatesStatus.textContent = statusText;
  }

  function syncWarehouseRateInputs() {
    dom.warehouseRateInputs.querySelectorAll('[data-warehouse-rate]').forEach((input) => {
      input.value = String(state.warehouse.rates[input.dataset.warehouseRate] ?? 0);
    });
  }

  function persistWarehouseRates() {
    if (dom.warehouseRatesStatus) dom.warehouseRatesStatus.textContent = 'Session only · export JSON or save an encrypted project to keep';
  }

  async function loadPrivateWarehouseRateFile() {
    try {
      const response = await fetch(WAREHOUSE_RATE_PRIVATE_FILE, { cache: 'no-store' });
      if (!response.ok) return false;
      const envelope = await response.json();
      applyWarehouseRates(decodeWarehouseRateEnvelope(envelope), 'Loaded private Caesar-14 quote file · not stored in browser');
      return true;
    } catch (error) {
      console.info('No local private warehouse-rate file was loaded.', error);
      return false;
    }
  }

  function exportWarehouseRates() {
    try {
      const envelope = createWarehouseRateEnvelope();
      const blob = new Blob([JSON.stringify(envelope, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = 'warehouse-rates.caesar14.json';
      document.body.append(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      showToast('Caesar-14 warehouse-rate JSON exported.');
    } catch (error) {
      console.error(error);
      showToast('The warehouse rates could not be exported.', true);
    }
  }

  async function importWarehouseRates(file) {
    try {
      if (file.size > 1024 * 1024) throw new Error('The warehouse-rate JSON file is unexpectedly large.');
      const envelope = JSON.parse(await file.text());
      const rates = decodeWarehouseRateEnvelope(envelope);
      applyWarehouseRates(rates, 'Imported for this session · export JSON or save an encrypted project to keep');
      refreshViews();
      showToast('Warehouse rates imported for this session.');
    } catch (error) {
      console.error(error);
      showToast(error.message || 'The warehouse-rate JSON could not be imported.', true);
    }
  }

  function zoomWarehousePreview(delta) {
    state.warehousePreviewTransform.scale = clamp(state.warehousePreviewTransform.scale + delta, 0.55, 3);
    applyWarehousePreviewTransform();
  }

  function resetWarehousePreview() {
    state.warehousePreviewTransform = { scale: 1, x: 0, y: 0 };
    applyWarehousePreviewTransform();
  }

  function applyWarehousePreviewTransform() {
    const scene = dom.warehousePreview.querySelector('.warehouse-scene');
    if (!scene) return;
    const { scale, x, y } = state.warehousePreviewTransform;
    scene.setAttribute('transform', `translate(${x.toFixed(2)} ${y.toFixed(2)}) translate(380 210) scale(${scale.toFixed(3)}) translate(-380 -210)`);
    const readout = dom.warehousePreview.querySelector('.preview-zoom-readout');
    if (readout) readout.textContent = `${Math.round(scale * 100)}%`;
  }

  function calculateWarehouseModel() {
    const warehouse = state.warehouse;
    const rates = warehouse.rates;
    const inventoryUnits = state.consolidated.reduce((sum, item) => sum + Math.max(0, item.maxQuantity || 0), 0);
    const unitsPerShelf = Math.max(1, warehouse.binsPerShelf * warehouse.unitsPerBin);
    warehouse.unitsPerShelf = unitsPerShelf;
    const storage = Core.planWarehouseStorage({
      inventoryUnits,
      shelfEnabled: warehouse.shelfEnabled,
      drawerEnabled: warehouse.drawerEnabled,
      palletEnabled: warehouse.palletEnabled,
      plannedPallets: warehouse.plannedPallets,
      shelfShare: warehouse.shelfShare,
      drawerShare: warehouse.drawerShare,
      unitsPerShelf,
      unitsPerDrawer: warehouse.unitsPerDrawer,
      unitsPerPallet: warehouse.unitsPerPallet,
    });
    const { shelfEnabled, drawerEnabled, palletEnabled, shelfUnits, drawerUnits, palletUnits, shelfShare, drawerShare, palletShare, shelfLocations, drawerLocations, pallets } = storage;
    const racks = shelfLocations > 0 ? Math.ceil(shelfLocations / warehouse.shelvesPerRack) : 0;
    const totalBins = shelfLocations * warehouse.binsPerShelf;
    const palletRateKey = `${warehouse.palletType}${warehouse.palletHeight}`;
    const palletRate = rates[palletRateKey] || 0;
    const outputCurrency = state.config.outputCurrency;
    const nokToOutput = outputCurrency === 'NOK' ? 1 : state.rates.NOK ?? null;

    const storageShelf = shelfLocations * rates.shelf;
    const storageDrawer = drawerLocations * rates.drawer;
    const storagePallet = pallets * palletRate;
    const receiving = warehouse.receipts * rates.receiptBase
      + warehouse.receipts * warehouse.receiptLines * rates.receiptLine;
    const orderHandling = warehouse.orders * rates.orderBase
      + warehouse.orders * warehouse.orderLines * rates.orderLine;
    const edi = warehouse.ediLabels * rates.edi;
    const freight = (warehouse.businessParcels + warehouse.privateParcels) * rates.parcel
      + warehouse.privateParcels * rates.privateSurcharge;
    const wms = warehouse.includeWms ? rates.wms : 0;
    const costs = [
      { label: 'Shelf storage', calculation: `${shelfLocations} locations × ${formatNok(rates.shelf)}`, value: storageShelf },
      { label: 'Drawer storage', calculation: `${drawerLocations} drawers × ${formatNok(rates.drawer)}`, value: storageDrawer },
      { label: `${warehouse.palletType === 'eu' ? 'EU' : 'Sea'} pallet storage`, calculation: `${pallets} pallets × ${formatNok(palletRate)}`, value: storagePallet },
      { label: 'Receiving', calculation: `${formatNumber(warehouse.receipts, 0)} receipts × (${formatNok(rates.receiptBase)} + ${formatNumber(warehouse.receiptLines, 0)} lines × ${formatNok(rates.receiptLine)})`, value: receiving },
      { label: 'Order handling', calculation: `${formatNumber(warehouse.orders, 0)} orders × (${formatNok(rates.orderBase)} + ${formatNumber(warehouse.orderLines, 0)} lines × ${formatNok(rates.orderLine)})`, value: orderHandling },
      { label: 'EDI labels', calculation: `${formatNumber(warehouse.ediLabels, 0)} labels × ${formatNok(rates.edi)}`, value: edi },
      { label: 'Outbound freight', calculation: `${formatNumber(warehouse.businessParcels + warehouse.privateParcels, 0)} parcels × ${formatNok(rates.parcel)} + ${formatNumber(warehouse.privateParcels, 0)} private surcharges`, value: freight },
      { label: 'Packaging', calculation: 'Monthly usage assumption', value: warehouse.packaging },
      { label: 'WMS license', calculation: warehouse.includeWms ? 'Included' : 'Excluded', value: wms },
    ];
    costs.forEach((item) => { item.converted = nokToOutput === null ? null : item.value * nokToOutput; });
    const monthlyNok = costs.reduce((sum, item) => sum + item.value, 0);
    const monthlyTotal = nokToOutput === null ? null : monthlyNok * nokToOutput;
    const { capacity, capacityShortfall } = storage;

    return {
      storageNok: storageShelf + storageDrawer + storagePallet,
      receivingNok: receiving,
      orderHandlingNok: orderHandling,
      ediNok: edi,
      freightNok: freight,
      packagingNok: warehouse.packaging,
      wmsNok: wms,
      inventoryUnits,
      distinctParts: state.consolidated.length,
      shelfShare,
      drawerShare,
      palletShare,
      shelfEnabled,
      drawerEnabled,
      palletEnabled,
      shelfUnits,
      drawerUnits,
      palletUnits,
      shelfLocations,
      drawerLocations,
      pallets,
      racks,
      totalBins,
      unitsPerShelf,
      palletRate,
      nokToOutput,
      outputCurrency,
      capacity,
      capacityShortfall,
      costs,
      monthlyNok,
      monthlyTotal,
      annualNok: monthlyNok * 12,
      annualTotal: monthlyTotal === null ? null : monthlyTotal * 12,
    };
  }

  function renderWarehouse() {
    if (!dom.warehouseView) return;
    const model = calculateWarehouseModel();
    const locationCount = model.shelfLocations + model.drawerLocations + model.pallets;
    const quoteScope = state.warehouse.country === '752' ? ' The current NOK rates come from the Norwegian 3PL quote; replace them with Swedish warehouse rates before making a decision.' : '';
    dom.warehouseDataNote.textContent = `${model.distinctParts} part numbers · ${formatNumber(model.inventoryUnits, 2)} units from workbook quantities. Storage choices and activity values marked “Assumed” are planning inputs.${quoteScope}`;
    dom.warehousePalletShare.textContent = `${model.pallets} planned pallet${model.pallets === 1 ? '' : 's'}`;
    dom.whUnitsShelf.value = String(model.unitsPerShelf);
    dom.whPlannedPallets.disabled = !state.warehouse.palletEnabled;
    dom.whPalletType.disabled = !state.warehouse.palletEnabled;
    dom.whPalletHeight.disabled = !state.warehouse.palletEnabled;
    dom.whShelfShare.disabled = !state.warehouse.shelfEnabled || !state.warehouse.drawerEnabled;
    dom.whDrawerShare.disabled = !state.warehouse.shelfEnabled || !state.warehouse.drawerEnabled;
    dom.warehouseShareWarning.textContent = model.capacityShortfall > 0
      ? `Selected storage is short by ${formatNumber(model.capacityShortfall, 0)} units. Enable shelf or drawer storage, or add pallet positions.`
      : state.warehouse.shareMessage || '';
    dom.warehouseCapacityLabel.textContent = `${formatNumber(model.capacity, 0)} unit capacity`;
    dom.warehouseSummaryCards.replaceChildren(
      summaryCard('Inventory units', formatNumber(model.inventoryUnits, 2)),
      summaryCard('Storage locations', formatNumber(locationCount, 0)),
      summaryCard(`Monthly estimate (${model.outputCurrency})`, formatMoney(model.monthlyTotal, model.outputCurrency)),
      summaryCard(`Annual estimate (${model.outputCurrency})`, formatMoney(model.annualTotal, model.outputCurrency)),
    );
    const fxText = model.nokToOutput === null ? 'Rate unavailable' : `1 NOK = ${formatNumber(model.nokToOutput, 6)} ${model.outputCurrency}`;
    dom.warehouseCurrencyReference.innerHTML = `<span>Quoted currency<strong>NOK</strong></span><span>Live reference<strong>${fxText}</strong></span><span>Displayed totals<strong>${model.outputCurrency}</strong></span>`;
    dom.warehouseOutputHeader.textContent = `${model.outputCurrency} / month`;
    dom.warehouseRateInputs.querySelectorAll('[data-warehouse-rate]').forEach((input) => {
      let reference = input.closest('label').querySelector('.rate-conversion');
      if (!reference) {
        reference = document.createElement('small');
        reference.className = 'rate-conversion';
        input.closest('label').append(reference);
      }
      const valueNok = Math.max(0, toNumber(input.value) ?? 0);
      reference.textContent = model.nokToOutput === null ? 'SEK rate unavailable' : `≈ ${formatMoney(valueNok * model.nokToOutput, model.outputCurrency)}`;
    });
    dom.warehousePreview.innerHTML = buildWarehouseIsoSvg(model);
    dom.warehouseStorageLegend.innerHTML = [
      `<span><i class="shelf"></i>${model.shelfLocations} shelf locations</span>`,
      `<span><i class="shelf"></i>${model.totalBins} bins across ${model.racks} racks</span>`,
      `<span><i class="drawer"></i>${model.drawerLocations} drawers</span>`,
      `<span><i class="pallet"></i>${model.pallets} ${state.warehouse.palletType === 'eu' ? 'EU' : 'sea'} pallets ≤${state.warehouse.palletHeight} cm</span>`,
    ].join('');
    dom.warehouseMonthlyTotal.textContent = formatMoney(model.monthlyTotal, model.outputCurrency);

    const rows = document.createDocumentFragment();
    model.costs.forEach((item) => {
      const row = document.createElement('tr');
      [item.label, item.calculation, formatNok(item.value), formatMoney(item.converted, model.outputCurrency)].forEach((value) => {
        const cell = document.createElement('td');
        cell.textContent = value;
        row.append(cell);
      });
      rows.append(row);
    });
    const totalRow = document.createElement('tr');
    totalRow.className = 'total-row';
    ['Monthly total', 'Sum of quoted and entered costs', formatNok(model.monthlyNok), formatMoney(model.monthlyTotal, model.outputCurrency)].forEach((value) => {
      const cell = document.createElement('td');
      cell.textContent = value;
      totalRow.append(cell);
    });
    rows.append(totalRow);
    dom.warehouseCostBody.replaceChildren(rows);
  }

  function buildWarehouseIsoSvg(model) {
    const elements = [];
    const project = (x, y, z) => [80 + (x + y) * 0.866, 205 + (x - y) * 0.31 - z];
    const points = (corners) => corners.map(([x, y, z]) => project(x, y, z).map((n) => n.toFixed(1)).join(',')).join(' ');
    const polygon = (corners, fill, stroke = '#365443', opacity = 1) => {
      elements.push(`<polygon points="${points(corners)}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="1" vector-effect="non-scaling-stroke"/>`);
    };
    const line = (a, b, color, width = 2, className = '') => {
      const [x1, y1] = project(...a);
      const [x2, y2] = project(...b);
      const classAttribute = className ? ` class="${className}"` : '';
      elements.push(`<line${classAttribute} x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="${color}" stroke-width="${width}" stroke-linecap="square" vector-effect="non-scaling-stroke"/>`);
    };
    const box = (x, y, z, width, depth, height, front, side, top, opacity = 1) => {
      polygon([[x, y, z], [x + width, y, z], [x + width, y, z + height], [x, y, z + height]], front, '#375246', opacity);
      polygon([[x + width, y, z], [x + width, y + depth, z], [x + width, y + depth, z + height], [x + width, y, z + height]], side, '#375246', opacity);
      polygon([[x, y, z + height], [x + width, y, z + height], [x + width, y + depth, z + height], [x, y + depth, z + height]], top, '#375246', opacity);
    };
    const label = (x, y, text, anchor = 'middle') => {
      elements.push(`<text x="${x}" y="${y}" text-anchor="${anchor}" fill="#294437" font-family="Inter,system-ui,sans-serif" font-size="12" font-weight="700">${text}</text>`);
    };

    polygon([[0, 0, 0], [600, 0, 0], [600, 245, 0], [0, 245, 0]], '#e5ece7', '#bccbc1', 0.72);
    for (let grid = 60; grid < 600; grid += 60) line([grid, 0, 0.2], [grid, 245, 0.2], '#c9d5cd', 0.7);
    for (let grid = 60; grid < 245; grid += 60) line([0, grid, 0.2], [600, grid, 0.2], '#c9d5cd', 0.7);

    const rackX = 25;
    const rackY = 35;
    const visibleRacks = Math.min(3, model.racks);
    const rackBays = visibleRacks;
    const rackLevels = Math.min(6, Math.max(1, state.warehouse.shelvesPerRack));
    const visibleBinsPerShelf = Math.min(8, Math.max(1, state.warehouse.binsPerShelf));
    const visibleShelfLocations = Math.min(model.shelfLocations, rackBays * rackLevels);
    const bayWidth = 70;
    const levelHeight = 34;
    const rackWidth = rackBays * bayWidth;
    const rackHeight = 9 + (rackLevels - 1) * levelHeight;
    let shelfCell = 0;
    if (rackBays > 0) {
      // Paint the rear uprights first so the shelves can sit in front of them.
      for (let post = 0; post <= rackBays; post += 1) {
        line([rackX + post * bayWidth, rackY + 40, 0], [rackX + post * bayWidth, rackY + 40, rackHeight], '#3b5f8a', 2.5, 'rack-upright rack-upright-rear');
      }
      for (let level = 0; level < rackLevels; level += 1) {
        const z = 5 + level * levelHeight;
        box(rackX, rackY, z, rackWidth, 40, 4, '#d9581f', '#af4319', '#ec7440');
        for (let bay = 0; bay < rackBays; bay += 1) {
          if (shelfCell >= visibleShelfLocations) continue;
          const binGap = 2;
          const innerWidth = bayWidth - 10;
          const binWidth = (innerWidth - binGap * (visibleBinsPerShelf - 1)) / visibleBinsPerShelf;
          for (let bin = 0; bin < visibleBinsPerShelf; bin += 1) {
            box(rackX + bay * bayWidth + 5 + bin * (binWidth + binGap), rackY + 5, z + 5, binWidth, 30, Math.min(24, levelHeight - 8), '#7b9b84', '#5d7e68', '#9bb6a2', 0.96);
          }
          shelfCell += 1;
        }
      }
      // Front uprights are intentionally painted after the shelf decks and bins.
      // In the isometric projection the right-most one is closest to the viewer.
      for (let post = 0; post <= rackBays; post += 1) {
        line([rackX + post * bayWidth, rackY, 0], [rackX + post * bayWidth, rackY, rackHeight], '#3b5f8a', 3.5, 'rack-upright rack-upright-front');
      }
      line([rackX, rackY, rackHeight], [rackX + rackWidth, rackY, rackHeight], '#3b5f8a', 3.5, 'rack-top-beam rack-top-beam-front');
      line([rackX, rackY + 40, rackHeight], [rackX + rackWidth, rackY + 40, rackHeight], '#3b5f8a', 2.5, 'rack-top-beam rack-top-beam-rear');
    }
    label(190, 38, `RACKS ${model.racks} · SHELVES ${model.shelfLocations} · BINS ${model.totalBins}`);

    const drawerX = 245;
    const drawerY = 58;
    const visibleDrawers = Math.min(7, model.drawerLocations);
    if (visibleDrawers > 0) {
      box(drawerX, drawerY, 0, 76, 46, visibleDrawers * 18 + 10, '#8fa99a', '#6c8f79', '#b9c9bf');
      for (let drawer = 0; drawer < visibleDrawers; drawer += 1) {
        const z = 6 + drawer * 18;
        polygon([[drawerX + 5, drawerY - 0.3, z], [drawerX + 71, drawerY - 0.3, z], [drawerX + 71, drawerY - 0.3, z + 14], [drawerX + 5, drawerY - 0.3, z + 14]], '#dfe7e1', '#587263');
        line([drawerX + 31, drawerY - 0.6, z + 7], [drawerX + 45, drawerY - 0.6, z + 7], '#587263', 1.6);
      }
    }
    label(385, 76, `DRAWERS · ${model.drawerLocations}`);

    const visiblePallets = Math.min(6, model.pallets);
    for (let pallet = 0; pallet < visiblePallets; pallet += 1) {
      const col = pallet % 3;
      const row = Math.floor(pallet / 3);
      const x = 375 + col * 66;
      const y = 75 + row * 82;
      box(x, y, 0, 52, 42, 6, '#a56f22', '#80551a', '#cf9b36');
      for (let slat = 5; slat < 50; slat += 11) box(x + slat, y + 3, 6, 6, 36, 3, '#c58c32', '#8c6022', '#e1ad55');
      const height = state.warehouse.palletHeight === '220' ? 62 : 38;
      box(x + 5, y + 5, 9, 42, 32, height, '#b9a176', '#8e7a57', '#d8c59f', 0.96);
      line([x + 26, y + 4.5, 10], [x + 26, y + 4.5, 9 + height], '#8e7a57', 1);
    }
    label(570, 126, `PALLETS · ${model.pallets}`);

    if (model.shelfLocations > visibleShelfLocations) label(235, 232, `+${model.shelfLocations - visibleShelfLocations} shelf locations`, 'start');
    if (model.drawerLocations > visibleDrawers) label(390, 238, `+${model.drawerLocations - visibleDrawers} drawers`, 'start');
    if (model.pallets > visiblePallets) label(545, 267, `+${model.pallets - visiblePallets} pallets`, 'start');

    const aria = `${model.shelfLocations} shelf locations, ${model.drawerLocations} drawers, and ${model.pallets} pallets, with capacity for ${formatNumber(model.capacity, 0)} units.`;
    const { scale, x, y } = state.warehousePreviewTransform;
    const transform = `translate(${x.toFixed(2)} ${y.toFixed(2)}) translate(380 210) scale(${scale.toFixed(3)}) translate(-380 -210)`;
    return `<div class="warehouse-preview-controls" aria-label="Warehouse preview controls"><button type="button" data-preview-action="out" aria-label="Zoom out">−</button><span class="preview-zoom-readout">${Math.round(scale * 100)}%</span><button type="button" data-preview-action="in" aria-label="Zoom in">+</button><button type="button" class="preview-reset" data-preview-action="reset">Reset</button></div><span class="warehouse-preview-hint">Drag to move · wheel to zoom</span><svg viewBox="0 0 760 420" role="img" aria-label="${aria}"><title>Estimated warehouse storage footprint</title><desc>${aria}</desc><g class="warehouse-scene" transform="${transform}">${elements.join('')}</g></svg>`;
  }

  function formatNok(value) {
    return `${formatNumber(value, 2)} NOK`;
  }

  function formatMoney(value, currency) {
    return Number.isFinite(value) ? `${formatNumber(value, 2)} ${currency}` : '—';
  }

  function buildWireModel() {
    return state.wireMode === 'underhood' ? buildUnderhoodWireModel() : buildOverviewWireModel();
  }

  function buildOverviewWireModel() {
    const warehouse = calculateWarehouseModel();
    const enabledSheets = state.workbook
      ? state.workbook.SheetNames.filter((name) => state.mappings[name]?.enabled)
      : [];
    const nodes = [];
    enabledSheets.forEach((name, index) => {
      nodes.push({ id: `sheet.${name}`, label: name, meta: 'Source sheet', detail: `${state.matrices[name]?.length || 0} imported rows`, category: 'source', x: 30, y: 35 + index * 68 });
    });
    nodes.push(
      { id: 'consolidated.parts', label: 'Part matching', meta: 'Normalize + group', detail: `${state.consolidated.length} part numbers`, category: 'calculation', x: 265, y: 35 },
      { id: 'consolidated.quantity', label: 'Maximum quantity', meta: 'Across source sheets', detail: `${formatNumber(warehouse.inventoryUnits, 2)} total units`, category: 'calculation', x: 265, y: 108 },
      { id: 'consolidated.price', label: 'Unit price', meta: 'Pricing source row', detail: 'Representative source price', category: 'calculation', x: 265, y: 181 },
      { id: 'consolidated.fx', label: 'Currency rate', meta: `Live to ${state.config.outputCurrency}`, detail: 'Daily reference rate', category: 'calculation', x: 265, y: 254 },
      { id: 'consolidated.discount', label: 'Discount + multiplier', meta: `${formatPercent(state.config.discount)} · ${formatNumber(state.config.multiplier, 2)}×`, detail: 'Sales price calculation', category: 'calculation', x: 265, y: 327 },
      { id: 'consolidated.shipping', label: 'Freight margin', meta: `${formatPercent(state.config.shippingMargin)} margin`, detail: 'Freight and import costs plus margin', category: 'calculation', x: 265, y: 400 },
      { id: 'consolidated.total', label: 'Line total', meta: `${state.config.outputCurrency} excl. VAT`, detail: 'Sales price + freight incl. margin', category: 'calculation', x: 265, y: 473 },
      { id: 'consolidated.freight', label: 'Freight & tolls', meta: `Kit + ${formatMoney(toOutput(state.freight.consolidatedShipment, state.freight.consolidatedCurrency), state.config.outputCurrency)} shipment`, detail: `Duty ${formatPercent(state.freight.dutyRate)}, insurance ${formatPercent(state.freight.insuranceRate)}, clearance fees`, category: 'assumption', x: 265, y: 546 },
      { id: 'consolidated.vat', label: 'VAT', meta: formatPercent(salesVatRate()), detail: 'Added to customer prices; import VAT is deductible', category: 'assumption', x: 265, y: 619 },
      { id: 'sales.groups', label: 'Sales groups', meta: `${state.groups.length} group${state.groups.length === 1 ? '' : 's'}`, detail: 'Expected sales and multiplier per group', category: 'assumption', x: 515, y: 473 },
      { id: 'sales.expected', label: 'Expected sales', meta: `${formatNumber(state.consolidated.reduce((sum, item) => sum + (item.salesPerYear || 0), 0), 0)} units / yr`, detail: 'Group, item or default sales assumption', category: 'calculation', x: 515, y: 546 },
      { id: 'output.dashboard', label: 'Profitability dashboard', meta: 'Net profit + payback', detail: 'Revenue, landed cost and warehouse costs', category: 'output', x: 765, y: 495 },
      { id: 'warehouse.inventory', label: 'Warehouse units', meta: 'From maximum quantities', detail: `${formatNumber(warehouse.inventoryUnits, 2)} units`, category: 'warehouse', x: 515, y: 35 },
      { id: 'warehouse.mix', label: 'Storage mix', meta: `${formatNumber(warehouse.shelfShare * 100, 0)}% shelf · ${formatNumber(warehouse.drawerShare * 100, 0)}% drawer`, detail: `${formatNumber(warehouse.palletShare * 100, 0)}% pallet`, category: 'warehouse', x: 515, y: 108 },
      { id: 'warehouse.bins', label: 'Bin capacity', meta: `${state.warehouse.binsPerShelf} bins × ${formatNumber(state.warehouse.unitsPerBin, 0)} units`, detail: `${state.warehouse.shelvesPerRack} shelf levels per rack`, category: 'warehouse', x: 515, y: 181 },
      { id: 'warehouse.locations', label: 'Storage locations', meta: `${warehouse.shelfLocations} shelf · ${warehouse.drawerLocations} drawer`, detail: `${warehouse.pallets} pallets · ${warehouse.racks} racks`, category: 'warehouse', x: 515, y: 254 },
      { id: 'warehouse.operations', label: 'Handling + freight', meta: `${formatNumber(state.warehouse.orders, 0)} orders / month`, detail: 'Receipts, lines, EDI and parcels', category: 'warehouse', x: 515, y: 327 },
      { id: 'warehouse.total', label: 'Warehouse total', meta: `${formatMoney(warehouse.monthlyTotal, warehouse.outputCurrency)} / month`, detail: `${formatMoney(warehouse.annualTotal, warehouse.outputCurrency)} / year`, category: 'warehouse', x: 515, y: 400 },
      { id: 'output.consolidated', label: 'Consolidated XLSX', meta: 'Readable workbook', detail: 'Parts, formulas and warehouse sheet', category: 'output', x: 765, y: 105 },
      { id: 'output.secure', label: 'Secure project', meta: 'Password-encrypted', detail: 'Workbook + settings + custom wires', category: 'output', x: 765, y: 235 },
      { id: 'output.preview', label: 'Isometric preview', meta: 'Live visual', detail: 'Racks, bins, drawers and pallets', category: 'output', x: 765, y: 365 },
    );

    const edges = [];
    enabledSheets.forEach((name) => edges.push({ from: `sheet.${name}`, to: 'consolidated.parts', label: 'parts + quantities' }));
    enabledSheets.forEach((name) => edges.push({ from: `sheet.${name}`, to: 'consolidated.freight', label: 'kit freight' }));
    edges.push(
      { from: 'consolidated.freight', to: 'consolidated.shipping', label: 'freight & import' },
      { from: 'consolidated.vat', to: 'consolidated.total', label: 'VAT column' },
      { from: 'consolidated.quantity', to: 'sales.expected', label: 'stock qty' },
      { from: 'sales.groups', to: 'sales.expected', label: 'group volumes' },
      { from: 'sales.expected', to: 'output.dashboard', label: 'units sold' },
      { from: 'consolidated.total', to: 'output.dashboard', label: 'unit prices' },
      { from: 'warehouse.total', to: 'output.dashboard', label: 'logistics cost' },
    );
    edges.push(
      { from: 'consolidated.parts', to: 'consolidated.quantity', label: 'group maximum' },
      { from: 'consolidated.parts', to: 'consolidated.price', label: 'pricing row' },
      { from: 'consolidated.price', to: 'consolidated.fx', label: 'convert' },
      { from: 'consolidated.fx', to: 'consolidated.discount', label: 'converted total' },
      { from: 'consolidated.discount', to: 'consolidated.total', label: 'sales price' },
      { from: 'consolidated.shipping', to: 'consolidated.total', label: 'freight' },
      { from: 'consolidated.quantity', to: 'warehouse.inventory', label: 'units' },
      { from: 'warehouse.inventory', to: 'warehouse.mix', label: 'allocate' },
      { from: 'warehouse.mix', to: 'warehouse.locations', label: 'shares' },
      { from: 'warehouse.bins', to: 'warehouse.locations', label: 'capacity' },
      { from: 'warehouse.locations', to: 'warehouse.total', label: 'storage cost' },
      { from: 'warehouse.operations', to: 'warehouse.total', label: 'monthly activity' },
      { from: 'consolidated.total', to: 'output.consolidated', label: 'parts output' },
      { from: 'warehouse.total', to: 'output.consolidated', label: 'warehouse sheet' },
      { from: 'consolidated.parts', to: 'output.secure', label: 'workbook data' },
      { from: 'warehouse.total', to: 'output.secure', label: 'cost assumptions' },
      { from: 'warehouse.locations', to: 'output.preview', label: 'geometry' },
    );
    const nodeIds = new Set(nodes.map((node) => node.id));
    state.wireRelations.forEach((edge) => {
      if (nodeIds.has(edge.from) && nodeIds.has(edge.to)) edges.push({ ...edge, custom: true });
    });
    return { nodes, edges, width: 980, nodeWidth: 180, nodeHeight: 50, showAllLabels: true };
  }

  function buildUnderhoodWireModel() {
    const warehouse = calculateWarehouseModel();
    const enabledSheets = state.workbook.SheetNames.filter((name) => state.mappings[name]?.enabled);
    const nodes = [];
    const edges = [];
    const sourceFields = [
      ['part', 'Part / SKU', 'column.part'],
      ['description', 'Description', 'column.description'],
      ['quantity', 'Quantity', 'column.quantity'],
      ['price', 'Unit price', 'column.unitPrice'],
      ['shipping', 'Shipping', 'column.shipping'],
    ];
    let sourceIndex = 0;
    enabledSheets.forEach((sheetName) => {
      const mapping = state.mappings[sheetName];
      sourceFields.forEach(([field, label, target]) => {
        const column = mapping[field];
        if (column === null || column === undefined) return;
        const id = `cell.${sheetName}.${field}`;
        nodes.push({ id, label: `${sheetName} · ${label}`, meta: `${columnName(column)} · source column`, detail: `${sheetName}!${columnName(column)} feeds the mapped ${label.toLowerCase()} column`, category: 'source', x: 20, y: 24 + sourceIndex * 43 });
        edges.push({ from: id, to: target, label: `${columnName(column)} → mapped` });
        if (field === 'part') edges.push({ from: id, to: 'column.sources', label: 'sheet identity' });
        sourceIndex += 1;
      });
    });

    nodes.push(
      { id: 'column.sources', label: 'Source sheets', meta: 'Consolidated column A', detail: `${enabledSheets.length} enabled source sheets`, category: 'calculation', x: 265, y: 24 },
      { id: 'column.part', label: 'Part', meta: 'Consolidated column B', detail: `${state.consolidated.length} normalized part numbers`, category: 'calculation', x: 265, y: 102 },
      { id: 'column.description', label: 'Description', meta: 'Consolidated column C', detail: 'Description from the representative row', category: 'calculation', x: 265, y: 180 },
      { id: 'column.quantity', label: 'Quantity', meta: 'Consolidated column D', detail: 'Maximum quantity across matching parts', category: 'calculation', x: 265, y: 258 },
      { id: 'column.unitPrice', label: 'Unit price USD', meta: 'Mapped price cell', detail: 'Source unit price used for calculations', category: 'calculation', x: 265, y: 336 },
      { id: 'column.shipping', label: 'Shipping', meta: 'Mapped freight cell', detail: 'Source freight amount before margin', category: 'calculation', x: 265, y: 414 },
      { id: 'formula.usdAmount', label: 'Amount in USD', meta: 'Quantity × unit price', detail: 'Source line amount normalized to USD', category: 'calculation', x: 510, y: 105 },
      { id: 'formula.usdSek', label: 'USD → SEK', meta: formatMoney(state.rates.USD, 'SEK'), detail: 'Live daily reference rate', category: 'assumption', x: 510, y: 183 },
      { id: 'formula.sekAmount', label: 'Amount in SEK', meta: 'USD amount × live rate', detail: 'Converted line amount before discount', category: 'calculation', x: 510, y: 261 },
      { id: 'formula.discount', label: `${formatPercent(state.config.discount)} discount`, meta: formatPercent(state.config.discount), detail: 'Configurable discount applied to SEK amount', category: 'assumption', x: 510, y: 339 },
      { id: 'formula.multiplier', label: 'Price multiplier', meta: `${formatNumber(state.config.multiplier, 2)}×`, detail: 'Applied after discount', category: 'assumption', x: 510, y: 417 },
      { id: 'formula.sales', label: 'Sales amount SEK', meta: 'Discounted × multiplier', detail: 'Calculated sales amount', category: 'calculation', x: 510, y: 495 },
      { id: 'formula.shippingMargin', label: 'Shipping margin', meta: formatPercent(state.config.shippingMargin), detail: 'Applied to the mapped shipping cost', category: 'assumption', x: 510, y: 573 },
      { id: 'formula.lineTotal', label: `Line total ${state.config.outputCurrency}`, meta: 'Sales + freight incl. margin', detail: 'Customer amount per consolidated part, excl. VAT', category: 'calculation', x: 510, y: 651 },
      { id: 'formula.kitFreight', label: 'Kit freight', meta: 'Per kit: sheet / per line / whole kit', detail: 'Kit freight setting applied to the pricing row', category: 'calculation', x: 510, y: 729 },
      { id: 'formula.consFreight', label: 'Consolidated shipment share', meta: `${formatMoney(toOutput(state.freight.consolidatedShipment, state.freight.consolidatedCurrency), state.config.outputCurrency)} by ${state.freight.allocation}`, detail: 'One shipment for the whole consolidated order, split across lines', category: 'assumption', x: 510, y: 807 },
      { id: 'formula.importCosts', label: 'Duty, insurance & fees', meta: `${formatPercent(state.freight.dutyRate)} duty · ${formatPercent(state.freight.insuranceRate)} ins.`, detail: 'Duty on purchase + freight + insurance, plus clearance fees', category: 'assumption', x: 510, y: 885 },
      { id: 'formula.vat', label: 'VAT', meta: formatPercent(salesVatRate()), detail: 'Line total incl. VAT for the customer', category: 'assumption', x: 510, y: 963 },
      { id: 'formula.landed', label: 'Landed cost', meta: 'Purchase + freight & import', detail: `Cost of the stock delivered to ${destinationConfig().name}`, category: 'calculation', x: 510, y: 1041 },
      { id: 'sales.groups', label: 'Sales groups', meta: `${state.groups.length} groups · ${Object.keys(state.items).length} item overrides`, detail: 'Group sales volumes and multipliers', category: 'assumption', x: 755, y: 804 },
      { id: 'sales.expected', label: 'Expected sales / yr', meta: `${formatNumber(state.consolidated.reduce((sum, item) => sum + (item.salesPerYear || 0), 0), 0)} units`, detail: 'Item override → group → default assumption', category: 'calculation', x: 755, y: 882 },
      { id: 'dashboard.net', label: 'Net profit / yr', meta: 'Revenue − landed cost − logistics', detail: 'Profitability dashboard result', category: 'calculation', x: 755, y: 960 },
      { id: 'output.dashboard', label: 'Dashboard', meta: 'Charts + item table', detail: 'Profit bridge, payback and group profit', category: 'output', x: 1000, y: 882 },
      { id: 'warehouse.inventory', label: 'Inventory units', meta: `${formatNumber(warehouse.inventoryUnits, 0)} units`, detail: 'Sum of consolidated maximum quantities', category: 'warehouse', x: 755, y: 102 },
      { id: 'warehouse.ratesNok', label: 'Quoted rates NOK', meta: 'Private workbook data', detail: 'Pallet, shelf, handling and freight rates', category: 'warehouse', x: 755, y: 180 },
      { id: 'warehouse.nokSek', label: 'NOK → SEK', meta: warehouse.nokToOutput === null ? 'Rate unavailable' : formatNumber(warehouse.nokToOutput, 6), detail: 'Live reference conversion for warehouse costs', category: 'assumption', x: 755, y: 258 },
      { id: 'warehouse.storageMix', label: 'Storage mix', meta: `${formatNumber(warehouse.shelfShare * 100, 0)} / ${formatNumber(warehouse.drawerShare * 100, 0)} / ${formatNumber(warehouse.palletShare * 100, 0)}%`, detail: 'Shelf, drawer and pallet allocation', category: 'warehouse', x: 755, y: 336 },
      { id: 'warehouse.binCapacity', label: 'Bin + rack capacity', meta: `${state.warehouse.binsPerShelf} bins · ${state.warehouse.shelvesPerRack} levels`, detail: `${formatNumber(state.warehouse.unitsPerBin, 0)} units per bin`, category: 'warehouse', x: 755, y: 414 },
      { id: 'warehouse.locations', label: 'Storage locations', meta: `${warehouse.shelfLocations} shelf · ${warehouse.drawerLocations} drawer`, detail: `${warehouse.pallets} pallets across ${warehouse.racks} racks`, category: 'warehouse', x: 755, y: 492 },
      { id: 'warehouse.activity', label: 'Monthly activity', meta: `${formatNumber(state.warehouse.orders, 0)} orders`, detail: 'Receipts, item lines, EDI and parcels', category: 'warehouse', x: 755, y: 570 },
      { id: 'warehouse.totalNok', label: 'Warehouse total NOK', meta: formatNok(warehouse.monthlyNok), detail: 'Monthly quoted and entered costs', category: 'warehouse', x: 755, y: 648 },
      { id: 'warehouse.totalSek', label: 'Warehouse total SEK', meta: formatMoney(warehouse.monthlyTotal, warehouse.outputCurrency), detail: 'Monthly total after live NOK conversion', category: 'warehouse', x: 755, y: 726 },
      { id: 'output.consolidated', label: 'Consolidated XLSX', meta: 'Readable workbook', detail: 'Parts, formulas, NOK rates and SEK references', category: 'output', x: 1000, y: 310 },
      { id: 'output.secure', label: 'Encrypted project', meta: 'AES-256-GCM', detail: 'Workbook, private rates, settings and custom wires', category: 'output', x: 1000, y: 466 },
      { id: 'output.preview', label: 'Isometric preview', meta: 'Racks + bins', detail: 'Live geometry driven by capacity assumptions', category: 'output', x: 1000, y: 622 },
    );

    edges.push(
      { from: 'column.quantity', to: 'formula.usdAmount', label: 'quantity' },
      { from: 'column.unitPrice', to: 'formula.usdAmount', label: 'unit price' },
      { from: 'formula.usdAmount', to: 'formula.sekAmount', label: 'USD amount' },
      { from: 'formula.usdSek', to: 'formula.sekAmount', label: 'FX rate' },
      { from: 'formula.sekAmount', to: 'formula.sales', label: 'base amount' },
      { from: 'formula.discount', to: 'formula.sales', label: 'discount' },
      { from: 'formula.multiplier', to: 'formula.sales', label: 'multiplier' },
      { from: 'column.shipping', to: 'formula.kitFreight', label: 'Frakt column' },
      { from: 'formula.kitFreight', to: 'formula.lineTotal', label: 'kit freight' },
      { from: 'formula.consFreight', to: 'formula.lineTotal', label: 'shipment share' },
      { from: 'formula.importCosts', to: 'formula.lineTotal', label: 'import costs' },
      { from: 'formula.shippingMargin', to: 'formula.lineTotal', label: 'margin' },
      { from: 'formula.lineTotal', to: 'formula.vat', label: 'excl. VAT' },
      { from: 'formula.sekAmount', to: 'formula.landed', label: 'purchase' },
      { from: 'formula.discount', to: 'formula.landed', label: 'discount' },
      { from: 'formula.kitFreight', to: 'formula.landed', label: 'kit freight' },
      { from: 'formula.consFreight', to: 'formula.landed', label: 'shipment share' },
      { from: 'formula.importCosts', to: 'formula.landed', label: 'import costs' },
      { from: 'column.quantity', to: 'sales.expected', label: 'stock qty' },
      { from: 'sales.groups', to: 'sales.expected', label: 'group volumes' },
      { from: 'sales.expected', to: 'dashboard.net', label: 'units sold' },
      { from: 'formula.lineTotal', to: 'dashboard.net', label: 'revenue' },
      { from: 'formula.landed', to: 'dashboard.net', label: 'cost of goods' },
      { from: 'warehouse.totalSek', to: 'dashboard.net', label: 'logistics' },
      { from: 'dashboard.net', to: 'output.dashboard', label: 'profitability' },
      { from: 'formula.sales', to: 'formula.lineTotal', label: 'sales amount' },
      { from: 'column.quantity', to: 'warehouse.inventory', label: 'max quantities' },
      { from: 'warehouse.inventory', to: 'warehouse.locations', label: 'units' },
      { from: 'warehouse.storageMix', to: 'warehouse.locations', label: 'allocation' },
      { from: 'warehouse.binCapacity', to: 'warehouse.locations', label: 'capacity' },
      { from: 'warehouse.locations', to: 'warehouse.totalNok', label: 'storage count' },
      { from: 'warehouse.activity', to: 'warehouse.totalNok', label: 'monthly operations' },
      { from: 'warehouse.ratesNok', to: 'warehouse.totalNok', label: 'NOK rates' },
      { from: 'warehouse.totalNok', to: 'warehouse.totalSek', label: 'NOK total' },
      { from: 'warehouse.nokSek', to: 'warehouse.totalSek', label: 'live FX' },
      { from: 'formula.lineTotal', to: 'output.consolidated', label: 'parts output' },
      { from: 'warehouse.totalSek', to: 'output.consolidated', label: 'warehouse sheet' },
      { from: 'column.sources', to: 'output.secure', label: 'source workbook' },
      { from: 'warehouse.ratesNok', to: 'output.secure', label: 'private rates' },
      { from: 'warehouse.locations', to: 'output.preview', label: 'geometry' },
    );
    const nodeIds = new Set(nodes.map((node) => node.id));
    state.wireRelations.forEach((edge) => {
      if (nodeIds.has(edge.from) && nodeIds.has(edge.to)) edges.push({ ...edge, custom: true });
    });
    return { nodes, edges, width: 1215, nodeWidth: 190, nodeHeight: 42, showAllLabels: false };
  }

  function renderWireView() {
    if (!dom.wireCanvas || !state.workbook) return;
    const model = applySavedWirePositions(buildWireModel());
    if (!model.nodes.some((node) => node.id === state.selectedWireNode)) state.selectedWireNode = model.nodes[0]?.id || null;
    const selected = state.selectedWireNode;
    const nodeById = new Map(model.nodes.map((node) => [node.id, node]));
    dom.wireView.classList.toggle('underhood-mode', state.wireMode === 'underhood');
    dom.wireDetailMode.value = state.wireMode;
    const width = model.width || 980;
    const nodeW = model.nodeWidth || 180;
    const nodeH = model.nodeHeight || 50;
    const height = Math.max(560, 55 + nodeH + Math.max(...model.nodes.map((node) => node.y)));
    const outgoingPalette = ['#d47b2a', '#c04474', '#725db0', '#218878', '#a77d12', '#ad4b3f', '#4676b5', '#8b5a2b'];
    const selectedOutgoing = model.edges.filter((edge) => edge.from === selected);
    const paths = model.edges.map((edge, index) => {
      const from = nodeById.get(edge.from);
      const to = nodeById.get(edge.to);
      if (!from || !to) return '';
      const x1 = from.x + nodeW;
      const y1 = from.y + nodeH / 2;
      const x2 = to.x;
      const y2 = to.y + nodeH / 2;
      const directionSign = x2 >= x1 ? 1 : -1;
      const bend = Math.max(42, Math.abs(x2 - x1) * 0.45);
      const c1x = x1 + directionSign * bend;
      const c2x = x2 - directionSign * bend;
      const direction = edge.to === selected ? 'incoming' : edge.from === selected ? 'outgoing' : '';
      const outgoingIndex = direction === 'outgoing' ? selectedOutgoing.indexOf(edge) : -1;
      const outgoingColor = outgoingIndex >= 0 ? outgoingPalette[outgoingIndex % outgoingPalette.length] : '';
      const colorStyle = outgoingColor ? ` style="--wire-color:${outgoingColor}"` : '';
      const className = ['wire-path', edge.custom ? 'custom' : '', direction].filter(Boolean).join(' ');
      const arrowClassName = ['wire-mid-arrow', edge.custom ? 'custom' : '', direction].filter(Boolean).join(' ');
      const marker = direction === 'incoming' ? 'arrow-in' : direction === 'outgoing' ? `arrow-out-${outgoingIndex % outgoingPalette.length}` : edge.custom ? 'arrow-custom' : 'arrow-default';
      const labelX = (x1 + x2) / 2;
      const labelY = (y1 + y2) / 2 - 5 - (index % 2) * 6;
      const label = model.showAllLabels || direction || edge.custom ? `<text class="wire-label" x="${labelX}" y="${labelY}" text-anchor="middle">${escapeMarkup(edge.label || '')}</text>` : '';
      const arrowPositions = Math.hypot(x2 - x1, y2 - y1) > 260 ? [0.38, 0.66] : [0.52];
      const midArrows = arrowPositions.map((t) => {
        const inverse = 1 - t;
        const x = inverse ** 3 * x1 + 3 * inverse ** 2 * t * c1x + 3 * inverse * t ** 2 * c2x + t ** 3 * x2;
        const y = inverse ** 3 * y1 + 3 * inverse ** 2 * t * y1 + 3 * inverse * t ** 2 * y2 + t ** 3 * y2;
        const dx = 3 * inverse ** 2 * (c1x - x1) + 6 * inverse * t * (c2x - c1x) + 3 * t ** 2 * (x2 - c2x);
        const dy = 3 * inverse ** 2 * (y1 - y1) + 6 * inverse * t * (y2 - y1) + 3 * t ** 2 * (y2 - y2);
        const angle = Math.atan2(dy, dx) * 180 / Math.PI;
        return `<g class="${arrowClassName}"${colorStyle} transform="translate(${x.toFixed(2)} ${y.toFixed(2)}) rotate(${angle.toFixed(2)})"><path d="M -5 -4 L 5 0 L -5 4 Z"/></g>`;
      }).join('');
      return `<path class="${className}"${colorStyle} d="M ${x1} ${y1} C ${c1x} ${y1}, ${c2x} ${y2}, ${x2} ${y2}" marker-end="url(#${marker})"/>${midArrows}${label}`;
    }).join('');
    const nodes = model.nodes.map((node) => `<g class="wire-node ${node.category}${node.id === selected ? ' selected' : ''}" data-wire-node="${escapeMarkup(node.id)}" transform="translate(${node.x} ${node.y})" role="button" aria-label="${escapeMarkup(`${node.label}. ${node.detail}`)}"><rect width="${nodeW}" height="${nodeH}" rx="8"/><text x="12" y="21">${escapeMarkup(node.label)}</text><text class="wire-node-meta" x="12" y="38">${escapeMarkup(node.meta)}</text></g>`).join('');
    const outgoingMarkers = outgoingPalette.map((color, index) => `<marker id="arrow-out-${index}" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="${color}"/></marker>`).join('');
    dom.wireCanvas.innerHTML = `<svg viewBox="0 0 ${width} ${height}" style="min-width:${width}px" role="img" aria-label="Data relationship wire diagram"><defs><marker id="arrow-default" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#a8b5ad"/></marker><marker id="arrow-in" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#3478b7"/></marker>${outgoingMarkers}<marker id="arrow-custom" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#8266a4"/></marker></defs>${paths}${nodes}</svg>`;
    dom.wireCanvas.querySelectorAll('[data-wire-node]').forEach((node) => {
      node.addEventListener('click', () => {
        if (Date.now() < state.suppressWireClickUntil) return;
        state.selectedWireNode = node.dataset.wireNode;
        renderWireView();
      });
      node.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          state.selectedWireNode = node.dataset.wireNode;
          renderWireView();
        }
      });
      node.setAttribute('tabindex', '0');
    });
    renderWireInspector(model);
    populateWireSelectors(model.nodes);
    renderCustomWires(model.nodes);
  }

  function applySavedWirePositions(model) {
    const positions = state.wireNodePositions[state.wireMode] || {};
    const maxY = Math.max(560, 55 + (model.nodeHeight || 50) + Math.max(...model.nodes.map((node) => node.y)));
    model.nodes = model.nodes.map((node) => {
      const saved = positions[node.id];
      if (!saved) return node;
      return {
        ...node,
        ...Core.clampWirePosition(saved, {
          width: model.width || 980,
          height: Math.max(maxY, saved.y + (model.nodeHeight || 50)),
          nodeWidth: model.nodeWidth || 180,
          nodeHeight: model.nodeHeight || 50,
        }),
      };
    });
    return model;
  }

  function wirePointerPosition(event, svg) {
    const bounds = svg.getBoundingClientRect();
    const viewBox = svg.viewBox.baseVal;
    return {
      x: viewBox.x + (event.clientX - bounds.left) * viewBox.width / Math.max(1, bounds.width),
      y: viewBox.y + (event.clientY - bounds.top) * viewBox.height / Math.max(1, bounds.height),
    };
  }

  function startWireDrag(event) {
    if (event.button !== 0) return;
    const nodeElement = event.target.closest?.('[data-wire-node]');
    const svg = dom.wireCanvas.querySelector('svg');
    if (!nodeElement || !svg) return;
    const model = applySavedWirePositions(buildWireModel());
    const node = model.nodes.find((item) => item.id === nodeElement.dataset.wireNode);
    if (!node) return;
    const point = wirePointerPosition(event, svg);
    state.selectedWireNode = node.id;
    state.wireDrag = {
      pointerId: event.pointerId,
      nodeId: node.id,
      offsetX: point.x - node.x,
      offsetY: point.y - node.y,
      startX: point.x,
      startY: point.y,
      moved: false,
    };
    dom.wireCanvas.classList.add('dragging-node');
    dom.wireCanvas.setPointerCapture?.(event.pointerId);
    event.preventDefault();
  }

  function moveWireDrag(event) {
    const drag = state.wireDrag;
    const svg = dom.wireCanvas.querySelector('svg');
    if (!drag || drag.pointerId !== event.pointerId || !svg) return;
    const point = wirePointerPosition(event, svg);
    const model = buildWireModel();
    const viewBox = svg.viewBox.baseVal;
    const position = Core.clampWirePosition({ x: point.x - drag.offsetX, y: point.y - drag.offsetY }, {
      width: viewBox.width,
      height: viewBox.height,
      nodeWidth: model.nodeWidth || 180,
      nodeHeight: model.nodeHeight || 50,
    });
    state.wireNodePositions[state.wireMode][drag.nodeId] = position;
    if (Math.hypot(point.x - drag.startX, point.y - drag.startY) > 3) drag.moved = true;
    renderWireView();
    event.preventDefault();
  }

  function finishWireDrag(event) {
    const drag = state.wireDrag;
    if (!drag || drag.pointerId !== event.pointerId) return;
    if (drag.moved) state.suppressWireClickUntil = Date.now() + 250;
    dom.wireCanvas.releasePointerCapture?.(event.pointerId);
    dom.wireCanvas.classList.remove('dragging-node');
    state.wireDrag = null;
    saveSettingsSoon();
    renderWireView();
  }

  function resetWireLayout() {
    state.wireNodePositions[state.wireMode] = {};
    state.wireDrag = null;
    saveSettingsSoon();
    renderWireView();
    showToast(`${state.wireMode === 'underhood' ? 'Under-the-hood' : 'Overview'} wire layout reset.`);
  }

  function renderWireInspector(model) {
    const node = model.nodes.find((item) => item.id === state.selectedWireNode);
    if (!node) return;
    const incoming = model.edges.filter((edge) => edge.to === node.id);
    const outgoing = model.edges.filter((edge) => edge.from === node.id);
    const labels = new Map(model.nodes.map((item) => [item.id, item.label]));
    dom.wireSelectedTitle.textContent = node.label;
    dom.wireSelectedDetail.textContent = node.detail;
    const fragments = [];
    incoming.forEach((edge) => fragments.push(Object.assign(document.createElement('span'), { textContent: `← ${labels.get(edge.from)} · ${edge.label}` })));
    outgoing.forEach((edge) => fragments.push(Object.assign(document.createElement('span'), { textContent: `→ ${labels.get(edge.to)} · ${edge.label}` })));
    if (!fragments.length) fragments.push(Object.assign(document.createElement('span'), { textContent: 'No direct wires.' }));
    dom.wireSelectedLinks.replaceChildren(...fragments);
  }

  function populateWireSelectors(nodes) {
    const currentFrom = dom.wireFrom.value;
    const currentTo = dom.wireTo.value;
    const optionsFor = (select, selectedValue, fallbackValue) => {
      select.replaceChildren(...nodes.map((node) => {
        const option = document.createElement('option');
        option.value = node.id;
        option.textContent = node.label;
        return option;
      }));
      const nextValue = nodes.some((node) => node.id === selectedValue) ? selectedValue : fallbackValue;
      if (nodes.some((node) => node.id === nextValue)) select.value = nextValue;
    };
    optionsFor(dom.wireFrom, currentFrom || state.selectedWireNode, state.selectedWireNode || nodes[0]?.id);
    optionsFor(dom.wireTo, currentTo, state.wireMode === 'underhood' ? 'warehouse.totalSek' : 'warehouse.total');
  }

  function addWireRelation(event) {
    event.preventDefault();
    const from = dom.wireFrom.value;
    const to = dom.wireTo.value;
    const label = dom.wireLabel.value.trim() || 'custom relationship';
    if (!from || !to || from === to) {
      showToast('Choose two different data points for the wire.', true);
      return;
    }
    if (state.wireRelations.some((edge) => edge.from === from && edge.to === to && edge.label === label)) {
      showToast('That custom wire already exists.', true);
      return;
    }
    state.wireRelations.push({ id: `wire-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, from, to, label });
    dom.wireLabel.value = '';
    state.selectedWireNode = from;
    renderWireView();
    showToast('Custom relationship added.');
  }

  function renderCustomWires(nodes) {
    const labels = new Map(nodes.map((node) => [node.id, node.label]));
    if (!state.wireRelations.length) {
      const empty = document.createElement('p');
      empty.className = 'muted';
      empty.textContent = 'No custom relationships yet.';
      dom.wireCustomList.replaceChildren(empty);
      return;
    }
    const rows = state.wireRelations.map((edge) => {
      const row = document.createElement('div');
      row.className = 'wire-custom-row';
      const text = document.createElement('span');
      text.textContent = `${labels.get(edge.from) || edge.from} → ${labels.get(edge.to) || edge.to}: ${edge.label}`;
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'icon-button';
      remove.setAttribute('aria-label', `Remove ${edge.label}`);
      remove.textContent = '×';
      remove.addEventListener('click', () => {
        state.wireRelations = state.wireRelations.filter((item) => item.id !== edge.id);
        renderWireView();
      });
      row.append(text, remove);
      return row;
    });
    dom.wireCustomList.replaceChildren(...rows);
  }

  function escapeMarkup(value) {
    return String(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  }

  function uniqueWorkbookSheetName(workbook, requested) {
    const base = String(requested || 'Kit').replace(/[\\/?*\[\]:]/g, ' ').trim().slice(0, 31) || 'Kit';
    if (!workbook.SheetNames.includes(base)) return base;
    for (let index = 2; index < 1000; index += 1) {
      const suffix = ` (${index})`;
      const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
      if (!workbook.SheetNames.includes(candidate)) return candidate;
    }
    return `Kit ${Date.now()}`.slice(0, 31);
  }

  async function readWorkbookFile(file) {
    const data = await file.arrayBuffer();
    return XLSX.read(data, {
      type: 'array', cellFormula: true, cellStyles: true, cellNF: true, cellDates: true, bookDeps: true, xlfn: true,
    });
  }

  // The only multi-kit price list supported so far is Häny's spare-part catalogue; its kits ship from Bulgaria.
  const CATALOGUE_DEFAULTS = Object.freeze({ manufacturer: 'Häny', originCountry: '100', fxMargin: 0.4, freightPercent: 3 });

  // Sheets that hold several kits one after another (see Core.parseKitCatalogue).
  function catalogueSheetsOf(workbook) {
    return workbook.SheetNames
      .map((name) => ({ name, catalogue: Core.parseKitCatalogue(XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: null })) }))
      .filter((entry) => entry.catalogue);
  }

  // The encrypted bundle is offered only when it exists next to the page (it is absent on a plain local copy).
  async function detectVault() {
    if (!window.PartsListVault) return;
    try {
      const response = await fetch(VAULT_FILE, { cache: 'no-store' });
      if (!response.ok) return;
      state.vaultBytes = new Uint8Array(await response.arrayBuffer());
      dom.vaultForm.classList.remove('hidden');
    } catch (error) {
      console.info('No encrypted document bundle is available.', error);
    }
  }

  async function unlockVault(event) {
    event.preventDefault();
    const key = dom.vaultKey.value;
    if (!key || !state.vaultBytes) return;
    dom.vaultOpen.disabled = true;
    setStatus('Unlocking the saved documents…', 'busy');
    try {
      const documents = await window.PartsListVault.open(state.vaultBytes, key);
      dom.vaultKey.value = '';
      const isRates = (document) => /.json$/i.test(document.name);
      const lists = documents.filter((document) => !isRates(document)).map((document) => new File([document.bytes], document.name));
      if (!lists.length) throw new Error('The saved documents contain no parts lists.');
      await addKitWorkbooks(lists);
      const rates = documents.find(isRates);
      if (rates) await importWarehouseRates(new File([rates.bytes], rates.name, { type: 'application/json' }));
      showToast('Saved documents imported.');
    } catch (error) {
      console.error(error);
      setStatus(`Could not unlock the saved documents: ${error.message}`, 'error');
      showToast(error.message || 'The saved documents could not be opened.', true);
    } finally {
      dom.vaultOpen.disabled = false;
    }
  }

  async function openWorkbookFile(file) {
    let isCatalogue = false;
    try {
      isCatalogue = catalogueSheetsOf(await readWorkbookFile(file)).length > 0;
    } catch (error) {
      // An unreadable file is reported by the normal import below.
    }
    if (isCatalogue) await addKitWorkbooks([file], { fresh: true });
    else await importWorkbook(file);
  }

  async function addKitWorkbooks(files, options = {}) {
    if (!state.workbook && !options.fresh) {
      const startsAsCatalogue = catalogueSheetsOf(await readWorkbookFile(files[0])).length > 0;
      if (!startsAsCatalogue) {
        await importWorkbook(files[0]);
        if (files.length > 1) await addKitWorkbooks(files.slice(1));
        return;
      }
    }
    const fresh = options.fresh === true || !state.workbook;
    setStatus(files.length === 1 ? 'Adding 1 workbook…' : `Adding ${files.length} workbooks…`, 'busy');
    try {
      const merged = XLSX.utils.book_new();
      if (!fresh) {
        state.workbook.SheetNames.forEach((name) => {
          if (!isGeneratedSheet(state.workbook.Sheets[name])) XLSX.utils.book_append_sheet(merged, state.workbook.Sheets[name], uniqueWorkbookSheetName(merged, name));
        });
      }
      let added = 0;
      const skipped = [];
      const duplicates = [];
      const newKits = {};
      const newMakers = {};
      for (const file of files) {
        const incoming = await readWorkbookFile(file);
        const fileBase = String(file.name).replace(/\.[^.]+$/, '');
        const catalogues = catalogueSheetsOf(incoming);
        if (catalogues.length) {
          // A catalogue file is split into one kit sheet per kit; its other sheets are not parts lists we want.
          const { manufacturer, originCountry } = CATALOGUE_DEFAULTS;
          catalogues.forEach(({ catalogue }) => {
            catalogue.kits.forEach((kit) => {
              const baseName = Core.catalogueSheetName(kit.name, manufacturer);
              // Adding the same price list twice would make every part "common", so kits already present are left out.
              if (merged.SheetNames.includes(baseName)) {
                duplicates.push(kit.name);
                return;
              }
              const targetName = uniqueWorkbookSheetName(merged, baseName);
              XLSX.utils.book_append_sheet(merged, XLSX.utils.aoa_to_sheet(Core.catalogueSheetRows(kit, { manufacturer, currency: catalogue.currency })), targetName);
              newKits[targetName] = {
                ...newKitSettings(targetName),
                manufacturer,
                originCountry,
                // Net prices are used as given, with no further discount.
                discountRate: 0,
                defaultCurrency: CURRENCIES.includes(catalogue.currency) ? catalogue.currency : null,
                // Starting points that can be changed per manufacturer: a buffer on the exchange rate and a freight
                // add-on as a percentage of each item's value.
                fxMargin: CATALOGUE_DEFAULTS.fxMargin,
                shippingMode: 'percent',
                shippingAmount: CATALOGUE_DEFAULTS.freightPercent,
              };
              newMakers[manufacturer] = { ...MAKER_DEFAULTS, location: ORIGIN_PRESETS[originCountry]?.place || '', shipmentMode: 'own', shipmentCurrency: state.config.outputCurrency, clearanceCurrency: state.config.outputCurrency };
              added += 1;
            });
          });
          incoming.SheetNames.filter((name) => !catalogues.some((entry) => entry.name === name) && !isGeneratedSheet(incoming.Sheets[name])).forEach((name) => skipped.push(name));
          continue;
        }
        incoming.SheetNames.forEach((name) => {
          if (isGeneratedSheet(incoming.Sheets[name])) return;
          const fallback = `${fileBase} - ${name}`;
          const targetName = merged.SheetNames.includes(name) ? uniqueWorkbookSheetName(merged, fallback) : uniqueWorkbookSheetName(merged, name);
          XLSX.utils.book_append_sheet(merged, incoming.Sheets[name], targetName);
          // The manufacturer and country of an added kit are not known; start from the file name and let the user set them.
          const headerText = (XLSX.utils.sheet_to_json(incoming.Sheets[name], { header: 1, defval: '', range: 0 }).slice(0, 12) || []).flat().join(' ');
          const guessed = CURRENCY_HOME_COUNTRY[currencyFromText(headerText)];
          newKits[targetName] = {
            ...newKitSettings(targetName),
            manufacturer: fileBase.slice(0, 80) || targetName,
            originCountry: guessed || state.supplier.country,
            shippingMode: 'kitTotal',
            shippingAmount: 0,
          };
          // Another manufacturer ships on its own unless it is told to share the project's consolidated shipment.
          newMakers[fileBase.slice(0, 80) || targetName] = { ...MAKER_DEFAULTS, location: ORIGIN_PRESETS[guessed || state.supplier.country]?.place || '', shipmentMode: 'own', shipmentCurrency: state.config.outputCurrency, clearanceCurrency: state.config.outputCurrency };
          added += 1;
        });
      }
      if (!added) throw new Error(duplicates.length ? 'These kits are already in the project.' : 'No source sheets were found in the selected workbook.');
      if (!fresh) {
        const settingsName = uniqueWorkbookSheetName(merged, 'PartsList settings');
        XLSX.utils.book_append_sheet(merged, buildSettingsSheet(), settingsName);
        merged.Workbook = merged.Workbook || {};
        merged.Workbook.Sheets = merged.SheetNames.map((name) => ({ Hidden: name === settingsName ? 1 : 0 }));
      }
      const bytes = XLSX.write(merged, { type: 'array', bookType: 'xlsx', cellStyles: true, bookSST: true });
      const projectName = (!fresh && state.fileName) || `${String(files[0].name).replace(/\.[^.]+$/, '')}.xlsx`;
      const mergedFile = new File([bytes], projectName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const origin = ORIGIN_PRESETS[CATALOGUE_DEFAULTS.originCountry];
      await importWorkbook(mergedFile, {
        addedKits: added,
        displayName: projectName,
        kitSettings: newKits,
        makerSettings: newMakers,
        skippedSheets: skipped,
        duplicateKits: duplicates,
        // A project that starts from a catalogue is that manufacturer's project.
        supplier: fresh && Object.values(newKits).every((kit) => kit.manufacturer === CATALOGUE_DEFAULTS.manufacturer)
          ? { name: CATALOGUE_DEFAULTS.manufacturer, location: origin.place, country: CATALOGUE_DEFAULTS.originCountry, countryName: origin.name, lat: origin.lat, lon: origin.lon }
          : null,
      });
    } catch (error) {
      console.error(error);
      setStatus(`Could not add kits: ${error.message}`, 'error');
      showToast(error.message || 'Kit import failed.', true);
    }
  }

  async function importWorkbook(file, options = {}) {
    if (!window.XLSX || !FormulaEngine) {
      showToast('Spreadsheet libraries are unavailable.', true);
      return;
    }

    setStatus(`Reading ${file.name}…`, 'busy');
    try {
      const workbook = await readWorkbookFile(file);

      if (!workbook.SheetNames.length) throw new Error('The workbook contains no sheets.');

      if (state.secureDownloadUrl) URL.revokeObjectURL(state.secureDownloadUrl);
      state.secureDownloadUrl = null;
      dom.secureDownload.removeAttribute('href');
      dom.secureDownload.classList.add('hidden');

      state.workbook = workbook;
      state.fileName = options.displayName || file.name;
      state.matrices = {};
      state.mappings = {};
      state.variables = [];
      state.wireRelations = [];
      state.wireNodePositions = { overview: {}, underhood: {} };
      state.wireDrag = null;
      state.selectedWireNode = 'consolidated.quantity';
      state.rates = {};
      state.rateDates = {};
      state.pairRates = {};
      state.liveRateCells = [];
      state.selectedCell = null;
      const persistedRates = { ...state.warehouse.rates };
      state.warehouse = {
        ...WAREHOUSE_DEFAULTS,
        rates: { ...WAREHOUSE_RATE_DEFAULTS, ...persistedRates },
      };
      state.config.vatRate = 0;
      state.vat = { unlocked: false, byCountry: { ...VAT_DEFAULTS } };
      state.freight = { ...FREIGHT_DEFAULTS };
      state.kits = {};
      state.makers = {};
      state.groups = [];
      state.items = {};
      state.sales = { ...SALES_DEFAULTS };
      state.scenario = { salesDelta: 0, purchaseDelta: 0, inboundDelta: 0, warehouseDelta: 0 };
      state.supplier = { ...SUPPLIER_DEFAULTS };
      state.customer = { country: '578' };
      FLOW_OVERRIDE_KEYS.forEach((key) => { state.flowOverrides[key] = null; });
      state.flowUnlocked = false;
      state.selectedKeys = new Set();
      state.selectionAnchor = null;
      state.sheetSelection = { sheetName: null, rows: new Set(), anchor: null };
      state.consolidatedSearch = '';
      state.consolidatedFilter = 'all';
      state.pendingDisabledKits = null;
      state.pendingMappings = null;
      state.pendingVariables = null;
      state.pendingWireRelations = null;
      state.importWarnings = [];
      state.guideDismissed = true; // the guide floats over the list, so it opens from the header button instead of covering the data after every import
      state.projectDirty = false;
      state.projectSaveLabel = 'Workbook imported';
      const restoredWorkbook = hydrateExportedSettings(workbook);
      if (options.kitSettings) Object.assign(state.kits, options.kitSettings);
      if (options.makerSettings) Object.assign(state.makers, options.makerSettings);
      if (options.supplier) state.supplier = { ...options.supplier };

      const engineSheets = {};
      workbook.SheetNames.forEach((sheetName) => {
        const matrix = worksheetToFormulaMatrix(workbook.Sheets[sheetName], sheetName);
        state.matrices[sheetName] = matrix;
        engineSheets[sheetName] = prepareEngineMatrix(sheetName, workbook.Sheets[sheetName], matrix);
        state.mappings[sheetName] = detectMapping(matrix);
        if (isGeneratedSheet(workbook.Sheets[sheetName])) state.mappings[sheetName].enabled = false;
      });
      applyPendingMappings();
      applyPendingDisabledKits();

      state.hf = FormulaEngine.buildFromSheets(engineSheets, {
        licenseKey: 'gpl-v3',
        precisionRounding: 10,
        evaluateNullToZero: false,
        useArrayArithmetic: true,
        useColumnIndex: true,
        maxRows: Math.max(MAX_IMPORT_ROWS + 100, 10000),
        maxColumns: Math.max(MAX_IMPORT_COLS + 20, 200),
      });

      state.sheetIds = {};
      workbook.SheetNames.forEach((name) => { state.sheetIds[name] = state.hf.getSheetId(name); });
      importNamedExpressions(workbook);
      restorePendingWorkbookState();
      rebuildDependencyIndex();

      dom.workbookName.textContent = state.fileName;
      // Sheets that PartsList generated (Consolidated, Profitability, settings…) are not counted as source sheets.
      const sourceSheetCount = workbook.SheetNames.filter((name) => !isGeneratedSheet(workbook.Sheets[name])).length;
      dom.sheetCount.textContent = `${sourceSheetCount} sheet${sourceSheetCount === 1 ? '' : 's'}`;
      dom.emptyPanel.classList.add('hidden');
      dom.controls.classList.remove('hidden');
      dom.tabs.classList.remove('hidden');
      dom.welcome.classList.add('hidden');
      dom.exportButton.disabled = false;
      dom.addKitLabel.classList.remove('hidden');

      dom.consolidatedSearch.value = '';
      syncControlsFromState();
      persistWarehouseRates();
      renderMappings();
      renderVariables();
      renderGroups();
      renderTabs();
      rebuildConsolidation();
      // Honour a linked view (#dashboard, #warehouse, a sheet name) when the workbook finishes loading.
      showView(viewFromHash() || VIEW.consolidated);
      renderAllSelectionBars();
      saveSettingsSoon();

      const enabled = Object.values(state.mappings).filter((mapping) => mapping.enabled).length;
      const restoredNote = restoredWorkbook ? ' Saved settings were restored from the workbook.' : '';
      const warningNote = state.importWarnings.length ? ` Warning: ${state.importWarnings.join(' ')}` : '';
      const addedNote = options.addedKits ? ` Added ${options.addedKits} new kit sheet${options.addedKits === 1 ? '' : 's'}.` : '';
      const skippedNote = (options.skippedSheets?.length ? ` Skipped other sheets: ${options.skippedSheets.join(', ')}.` : '')
        + (options.duplicateKits?.length ? ` Already in the project, not added: ${options.duplicateKits.join(', ')}.` : '');
      setStatus(`Loaded ${state.fileName}. ${enabled} of ${sourceSheetCount} sheets are included in consolidation.${addedNote}${skippedNote}${restoredNote}${warningNote}`, state.importWarnings.length ? 'warning' : '');
      if (options.addedKits) showToast(`${options.addedKits} new kit sheet${options.addedKits === 1 ? '' : 's'} added. Review each kit's manufacturer, route, currency and discount.`);
      if (state.importWarnings.length) showToast(`Import warning: ${state.importWarnings.join(' ')}`);
      await refreshRates(false);
      markProjectSaved(restoredWorkbook ? 'Workbook settings restored' : 'Workbook imported');
      setLanguage(state.language, false);
      renderSetupGuide();
    } catch (error) {
      console.error(error);
      setStatus(`Could not import the workbook: ${error.message}`, 'error');
      showToast(error.message || 'Workbook import failed.', true);
    }
  }

  function isGeneratedSheet(sheet) {
    const marker = String(sheet?.A1?.v ?? '').trim().toLowerCase();
    return ['consolidated parts', 'warehouse cost estimate', 'assumption', 'profitability estimate', 'partslist project data', SETTINGS_SHEET_MARKER].includes(marker);
  }

  function readSettingsSheet(workbook) {
    const name = workbook.SheetNames.find((sheetName) => String(workbook.Sheets[sheetName]?.A1?.v ?? '').trim().toLowerCase() === SETTINGS_SHEET_MARKER);
    if (!name) return null;
    const sheet = workbook.Sheets[name];
    let json = '';
    for (let row = 3; row < 400; row += 1) {
      const value = sheet[`A${row}`]?.v;
      if (typeof value !== 'string' || !value) break;
      json += value;
    }
    try {
      return JSON.parse(json);
    } catch (error) {
      console.warn('The PartsList settings sheet could not be read.', error);
      return null;
    }
  }

  function hydrateExportedSettings(workbook) {
    hydrateWarehouseAndAssumptionSheets(workbook);
    const settings = readSettingsSheet(workbook);
    return settings ? applySettings(settings) : false;
  }

  function hydrateWarehouseAndAssumptionSheets(workbook) {
    const warehouseName = workbook.SheetNames.find((name) => String(workbook.Sheets[name]?.A1?.v ?? '').trim().toLowerCase() === 'warehouse cost estimate');
    if (warehouseName) {
      const sheet = workbook.Sheets[warehouseName];
      const assignNumber = (target, key, address, minimum = 0) => {
        const value = toNumber(sheet[address]?.v);
        if (Number.isFinite(value) && value >= minimum) target[key] = value;
      };
      const shelfShare = toNumber(sheet.B7?.v);
      const drawerShare = toNumber(sheet.B8?.v);
      if (Number.isFinite(shelfShare) && shelfShare >= 0 && shelfShare <= 1) state.warehouse.shelfShare = Math.round(shelfShare * 10000) / 100;
      if (Number.isFinite(drawerShare) && drawerShare >= 0 && drawerShare <= 1) state.warehouse.drawerShare = Math.round(drawerShare * 10000) / 100;
      if (Number.isFinite(shelfShare)) state.warehouse.shelfEnabled = shelfShare > 0;
      if (Number.isFinite(drawerShare)) state.warehouse.drawerEnabled = drawerShare > 0;
      assignNumber(state.warehouse, 'binsPerShelf', 'B10', 1);
      assignNumber(state.warehouse, 'shelvesPerRack', 'B11', 1);
      assignNumber(state.warehouse, 'unitsPerBin', 'B12', 1);
      assignNumber(state.warehouse, 'unitsPerDrawer', 'B14', 1);
      assignNumber(state.warehouse, 'unitsPerPallet', 'B15', 1);
      state.warehouse.unitsPerShelf = state.warehouse.binsPerShelf * state.warehouse.unitsPerBin;
      state.warehouse.palletType = String(sheet.B16?.v ?? '').toUpperCase() === 'SEA' ? 'sea' : 'eu';
      state.warehouse.palletHeight = String(sheet.B17?.v ?? '') === '220' ? '220' : '120';
      const rateCells = {
        eu120: 'E5', eu220: 'E6', sea120: 'E7', sea220: 'E8', shelf: 'E9', drawer: 'E10', edi: 'E11',
        receiptBase: 'E12', receiptLine: 'E13', orderBase: 'E14', orderLine: 'E15', parcel: 'E16', privateSurcharge: 'E17', wms: 'E18',
      };
      Object.entries(rateCells).forEach(([key, address]) => assignNumber(state.warehouse.rates, key, address));
      const activityCells = {
        receipts: 'B29', receiptLines: 'B30', orders: 'B31', orderLines: 'B32', ediLabels: 'B33',
        businessParcels: 'B34', privateParcels: 'B35', packaging: 'B36',
      };
      Object.entries(activityCells).forEach(([key, address]) => assignNumber(state.warehouse, key, address));
      const exportedPallets = toNumber(sheet.B25?.v);
      if (Number.isFinite(exportedPallets) && exportedPallets >= 0) {
        state.warehouse.plannedPallets = Math.round(exportedPallets);
        state.warehouse.palletEnabled = exportedPallets > 0;
      }
      state.warehouse.includeWms = sheet.B37?.v === true || String(sheet.B37?.v ?? '').toUpperCase() === 'TRUE';
    }

    const assumptionsName = workbook.SheetNames.find((name) => String(workbook.Sheets[name]?.A1?.v ?? '').trim().toLowerCase() === 'assumption');
    if (!assumptionsName) return;
    const sheet = workbook.Sheets[assumptionsName];
    const discount = toNumber(sheet.B2?.v);
    const multiplier = toNumber(sheet.B3?.v);
    const shippingMargin = toNumber(sheet.B4?.v);
    const outputCurrency = String(sheet.B5?.v ?? '').toUpperCase();
    if (Number.isFinite(discount) && discount >= 0 && discount <= 1) state.config.discount = discount;
    if (Number.isFinite(multiplier) && multiplier >= 0) state.config.multiplier = multiplier;
    if (Number.isFinite(shippingMargin) && shippingMargin >= 0 && shippingMargin < 1) state.config.shippingMargin = shippingMargin;
    if (CURRENCIES.includes(outputCurrency)) state.config.outputCurrency = outputCurrency;
  }

  function worksheetToFormulaMatrix(worksheet, sheetName = 'Worksheet') {
    if (!worksheet || !worksheet['!ref']) return [[]];
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const bounds = Core.importBounds(range.e.r, range.e.c, MAX_IMPORT_ROWS, MAX_IMPORT_COLS);
    const lastRow = bounds.lastRow;
    const lastCol = bounds.lastColumn;
    if (bounds.truncatedRows || bounds.truncatedColumns) {
      const imported = `${lastRow + 1} rows × ${lastCol + 1} columns`;
      const source = `${bounds.sourceRows} rows × ${bounds.sourceColumns} columns`;
      state.importWarnings.push(`“${sheetName}” was limited to ${imported} from ${source}.`);
    }
    const matrix = [];

    for (let row = 0; row <= lastRow; row += 1) {
      const values = [];
      for (let col = 0; col <= lastCol; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[address];
        if (!cell) {
          values.push(null);
        } else if (cell.f) {
          values.push(`=${cell.f}`);
        } else if (cell.v instanceof Date) {
          values.push(cell.v);
        } else if (typeof cell.v === 'string' && cell.v.startsWith('=')) {
          values.push(`'${cell.v}`);
        } else {
          values.push(cell.v ?? null);
        }
      }
      while (values.length && values[values.length - 1] === null) values.pop();
      matrix.push(values);
    }
    while (matrix.length > 1 && matrix[matrix.length - 1].length === 0) matrix.pop();
    return matrix.length ? matrix : [[]];
  }

  function prepareEngineMatrix(sheetName, worksheet, sourceMatrix) {
    return sourceMatrix.map((row, rowIndex) => row.map((value, colIndex) => {
      if (typeof value !== 'string' || !value.startsWith('=')) return value;
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const currencyPair = extractCurrencyPair(value);
      if (currencyPair) {
        const fallback = extractFormulaFallback(value);
        state.liveRateCells.push({ sheetName, row: rowIndex, col: colIndex, ...currencyPair });
        return fallback ?? (typeof worksheet[address]?.v === 'number' ? worksheet[address].v : 1);
      }

      const arrayArithmetic = value.match(/^=(\$?[A-Z]{1,3}\$?\d+):\$?[A-Z]{1,3}\$?\d+([*\/+\-].+)$/i);
      if (arrayArithmetic) return `=${arrayArithmetic[1]}${arrayArithmetic[2]}`;
      return value;
    }));
  }

  function extractCurrencyPair(formula) {
    const compact = String(formula).replace(/\s+/g, '').toUpperCase();
    const match = compact.match(/CURRENCY:([A-Z]{3})([A-Z]{3})/);
    return match ? { from: match[1], to: match[2] } : null;
  }

  function extractFormulaFallback(formula) {
    const match = String(formula).match(/,\s*(-?\d+(?:\.\d+)?)\s*\)+$/);
    return match ? Number(match[1]) : null;
  }

  function detectMapping(matrix) {
    let best = { row: 0, score: -1, fields: {} };
    const candidates = Math.min(matrix.length, 25);

    for (let row = 0; row < candidates; row += 1) {
      const fields = {};
      let score = 0;
      (matrix[row] || []).forEach((value, col) => {
        const header = normalizeHeader(value);
        if (!header) return;
        Object.entries(HEADER_TERMS).forEach(([field, terms]) => {
          if (fields[field] !== undefined) return;
          const matchScore = headerMatchScore(header, terms);
          if (matchScore > 0) {
            fields[field] = col;
            score += matchScore + (field === 'part' || field === 'quantity' ? 3 : 0);
          }
        });
      });
      const completeness = Object.keys(fields).length;
      score += completeness * 2;
      if (fields.part !== undefined && fields.quantity !== undefined) score += 12;
      if (score > best.score) best = { row, score, fields };
    }

    let dataStartRow = best.row + 1;
    if (best.fields.part !== undefined && best.fields.quantity !== undefined) {
      const detected = matrix.findIndex((row, rowIndex) => rowIndex > best.row
        && String(row?.[best.fields.part] ?? '').trim() !== ''
        && toNumber(row?.[best.fields.quantity]) !== null);
      if (detected >= 0) dataStartRow = detected;
    }

    const maxColumns = Math.max(0, ...matrix.slice(best.row, dataStartRow).map((row) => row?.length || 0));
    const combinedHeaders = Array.from({ length: maxColumns }, (_, col) => matrix
      .slice(best.row, dataStartRow)
      .map((row) => normalizeHeader(row?.[col]))
      .filter(Boolean)
      .join(' '));

    ['discountedTotal', 'sellingTotal', 'shipping', 'shippingWithMargin', 'currency'].forEach((field) => {
      let bestColumn = null;
      let bestScore = 0;
      combinedHeaders.forEach((header, col) => {
        const matchScore = headerMatchScore(header, HEADER_TERMS[field]);
        if (matchScore > bestScore) {
          bestScore = matchScore;
          bestColumn = col;
        }
      });
      if (bestColumn !== null) best.fields[field] = bestColumn;
    });

    return {
      enabled: best.fields.part !== undefined && best.fields.quantity !== undefined,
      headerRow: best.row,
      dataStartRow,
      ...Object.fromEntries(Object.keys(FIELD_LABELS).map((field) => [field, best.fields[field] ?? null])),
    };
  }

  function normalizeHeader(value) {
    if (value === null || value === undefined) return '';
    return String(value)
      .toLowerCase()
      .replace(/[_\-–—/\\()\[\].:%]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function headerMatchScore(header, terms) {
    let score = 0;
    terms.forEach((term) => {
      if (header === term) score = Math.max(score, 7);
      else if (header.startsWith(`${term} `) || header.endsWith(` ${term}`)) score = Math.max(score, 5);
      else if (term.length >= 5 && (` ${header} `).includes(` ${term} `)) score = Math.max(score, 3);
    });
    return score;
  }

  function renderMappings() {
    dom.mappingList.replaceChildren();
    state.workbook.SheetNames.forEach((sheetName) => {
      if (isGeneratedSheet(state.workbook.Sheets[sheetName])) return;
      const mapping = state.mappings[sheetName];
      const mappingReady = mapping.part !== null
        && mapping.quantity !== null
        && (mapping.price !== null || mapping.discountedTotal !== null || mapping.sellingTotal !== null);
      const card = document.createElement('article');
      card.className = `mapping-card${mapping.enabled ? '' : ' disabled'}${mappingReady ? '' : ' needs-review'}`;

      const head = document.createElement('div');
      head.className = 'mapping-card-head';
      const enabled = document.createElement('input');
      enabled.type = 'checkbox';
      enabled.checked = mapping.enabled;
      enabled.setAttribute('aria-label', `Include ${sheetName}`);
      const title = document.createElement('strong');
      title.textContent = sheetName;
      const hint = document.createElement('span');
      hint.className = 'mapping-hint';
      hint.textContent = `${state.matrices[sheetName].length} rows · ${mappingReady ? 'Ready' : 'Needs review'}`;
      head.append(enabled, title, hint);

      const fields = document.createElement('div');
      fields.className = 'mapping-fields';
      fields.append(createHeaderRowControl(sheetName, mapping, card));
      fields.append(createDataStartRowControl(sheetName, mapping));
      Object.keys(FIELD_LABELS).forEach((field) => fields.append(createColumnControl(sheetName, field, mapping)));

      enabled.addEventListener('change', () => {
        mapping.enabled = enabled.checked;
        card.classList.toggle('disabled', !mapping.enabled);
        saveSettingsSoon();
        rebuildConsolidation();
        refreshRates(false);
        renderMappings();
      });

      const details = document.createElement('details');
      details.className = 'mapping-details';
      details.open = mapping.enabled && !mappingReady;
      const summary = document.createElement('summary');
      summary.textContent = mappingReady ? 'Column mapping (advanced)' : 'Fix detected columns';
      details.append(summary, fields);

      card.append(head);
      if (!isGeneratedSheet(state.workbook.Sheets[sheetName])) card.append(createKitFreightControl(sheetName));
      card.append(details);
      dom.mappingList.append(card);
    });
  }

  function createHeaderRowControl(sheetName, mapping, card) {
    const label = document.createElement('label');
    label.textContent = 'Header row';
    const input = document.createElement('input');
    input.type = 'number';
    input.min = '1';
    input.max = String(Math.max(state.matrices[sheetName].length, 1));
    input.value = String(mapping.headerRow + 1);
    input.addEventListener('change', () => {
      mapping.headerRow = clamp((toNumber(input.value) ?? 1) - 1, 0, Math.max(0, state.matrices[sheetName].length - 1));
      const replacement = document.createElement('div');
      replacement.className = 'mapping-fields';
      replacement.append(createHeaderRowControl(sheetName, mapping, card));
      replacement.append(createDataStartRowControl(sheetName, mapping));
      Object.keys(FIELD_LABELS).forEach((field) => replacement.append(createColumnControl(sheetName, field, mapping)));
      card.querySelector('.mapping-fields').replaceWith(replacement);
      saveSettingsSoon();
      rebuildConsolidation();
      renderMappings();
    });
    label.append(input);
    return label;
  }

  function createDataStartRowControl(sheetName, mapping) {
    const label = document.createElement('label');
    label.textContent = 'First data row';
    const input = document.createElement('input');
    input.type = 'number';
    input.min = String(mapping.headerRow + 2);
    input.max = String(Math.max(state.matrices[sheetName].length, mapping.headerRow + 2));
    input.value = String(mapping.dataStartRow + 1);
    input.addEventListener('change', () => {
      mapping.dataStartRow = clamp((toNumber(input.value) ?? mapping.headerRow + 2) - 1, mapping.headerRow + 1, Math.max(mapping.headerRow + 1, state.matrices[sheetName].length - 1));
      saveSettingsSoon();
      rebuildConsolidation();
      renderMappings();
    });
    label.append(input);
    return label;
  }

  function createColumnControl(sheetName, field, mapping) {
    const label = document.createElement('label');
    label.textContent = FIELD_LABELS[field];
    const select = document.createElement('select');
    const none = document.createElement('option');
    none.value = '';
    none.textContent = 'Not mapped';
    select.append(none);
    const matrix = state.matrices[sheetName];
    const maxColumns = Math.max(0, ...matrix.slice(mapping.headerRow, mapping.dataStartRow).map((row) => row?.length || 0));
    Array.from({ length: maxColumns }).forEach((_, index) => {
      const header = matrix
        .slice(mapping.headerRow, mapping.dataStartRow)
        .map((row) => {
          const value = row?.[index];
          return typeof value === 'string' && value.startsWith('=') ? '' : displayValue(value);
        })
        .filter(Boolean)
        .join(' / ');
      const option = document.createElement('option');
      option.value = String(index);
      option.textContent = `${columnName(index)} — ${displayValue(header) || translateUi('(blank)')}`;
      option.translate = false;
      option.selected = mapping[field] === index;
      select.append(option);
    });
    select.value = mapping[field] === null ? '' : String(mapping[field]);
    select.addEventListener('change', () => {
      mapping[field] = select.value === '' ? null : Number(select.value);
      saveSettingsSoon();
      rebuildConsolidation();
      if (field === 'currency') refreshRates(false);
      renderMappings();
    });
    label.append(select);
    return label;
  }

  function rebuildConsolidation() {
    if (!state.workbook || !state.hf) return;
    const grouped = new Map();
    const kitOccurrences = new Map();

    state.workbook.SheetNames.forEach((sheetName) => {
      const mapping = state.mappings[sheetName];
      if (!mapping.enabled || mapping.part === null || mapping.quantity === null) return;
      const kit = kitSettings(sheetName);
      const matrix = state.matrices[sheetName];
      const priceHeader = matrix
        .slice(mapping.headerRow, mapping.dataStartRow)
        .map((row) => displayValue(row?.[mapping.price]))
        .filter(Boolean)
        .join(' ');
      const occurrences = [];

      for (let row = mapping.dataStartRow; row < matrix.length; row += 1) {
        const partValue = readCalculatedValue(sheetName, row, mapping.part);
        const part = partValue === null || partValue === undefined ? '' : String(partValue).trim();
        if (!part || /^(?:total(?:s)?|effective\b|confidential\b)/i.test(part)) continue;

        const quantity = toNumber(readCalculatedValue(sheetName, row, mapping.quantity));
        if (quantity === null) continue;
        const key = normalizePart(part);
        if (!key) continue;

        const currencyValue = mapping.currency === null ? null : readCalculatedValue(sheetName, row, mapping.currency);
        // Order: a currency column in the sheet, the kit's own setting, a currency named in the price header,
        // then the project default.
        const currency = normalizeCurrency(currencyValue)
          || kit.defaultCurrency
          || currencyFromText(priceHeader)
          || state.config.sourceCurrency;

        const occurrence = {
          key,
          part,
          description: mapping.description === null ? '' : displayValue(readCalculatedValue(sheetName, row, mapping.description)),
          altPart: mapping.altPart === null ? '' : cleanAltPart(displayValue(readCalculatedValue(sheetName, row, mapping.altPart))),
          quantity,
          price: mapping.price === null ? null : toNumber(readCalculatedValue(sheetName, row, mapping.price)),
          sourceDiscountedTotal: mapping.discountedTotal === null ? null : toNumber(readCalculatedValue(sheetName, row, mapping.discountedTotal)),
          sourceSellingTotal: mapping.sellingTotal === null ? null : toNumber(readCalculatedValue(sheetName, row, mapping.sellingTotal)),
          shipping: mapping.shipping === null ? null : toNumber(readCalculatedValue(sheetName, row, mapping.shipping)),
          sourceShippingWithMargin: mapping.shippingWithMargin === null ? null : toNumber(readCalculatedValue(sheetName, row, mapping.shippingWithMargin)),
          currency,
          fxMargin: kitFxMargin(kit),
          freightPercent: kit.shippingMode === 'percent' ? Math.max(0, kit.shippingAmount) / 100 : null,
          discountRate: kitDiscount(kit),
          dutyRate: kit.dutyRate,
          manufacturer: kit.manufacturer,
          originCountry: kit.originCountry,
          sheetName,
          row,
        };

        occurrences.push(occurrence);
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(occurrence);
      }
      kitOccurrences.set(sheetName, occurrences);
    });

    kitOccurrences.forEach((occurrences, sheetName) => {
      const kit = kitSettings(sheetName);
      const stockedLines = occurrences.filter((occurrence) => occurrence.quantity > 0).length;
      occurrences.forEach((occurrence) => {
        if (occurrence.freightPercent !== null) {
          // A percentage of the line's purchase value, already in the output currency.
          const fx = effectiveFx(occurrence.currency, occurrence.fxMargin);
          occurrence.kitFreight = occurrence.price === null || fx === null
            ? null
            : occurrence.price * fx * occurrence.quantity * (1 - occurrence.discountRate) * occurrence.freightPercent;
          return;
        }
        occurrence.kitFreight = toOutput(occurrenceFreight(occurrence, kit, stockedLines), kit.shippingCurrency);
      });
    });

    const items = [];
    grouped.forEach((occurrences) => items.push(buildConsolidatedItem(occurrences)));
    items.sort((a, b) => Number(b.common) - Number(a.common) || a.part.localeCompare(b.part, undefined, { numeric: true }));
    state.consolidated = items.filter((item) => !item.excluded);
    state.excludedItems = items.filter((item) => item.excluded);
    applyLandedCosts(state.consolidated, true);
    applyLandedCosts(state.excludedItems, false);
    state.kitSummaries = buildKitSummaries(kitOccurrences);

    const known = new Set(items.map((item) => item.key));
    [...state.selectedKeys].forEach((key) => { if (!known.has(key)) state.selectedKeys.delete(key); });
    refreshViews();
    renderSelectionBarCount();
  }

  // Freight for one source line, in the kit's freight currency.
  function occurrenceFreight(occurrence, kit, stockedLines) {
    if (kit.shippingMode === 'perLine') return kit.shippingAmount;
    if (kit.shippingMode === 'kitTotal') return occurrence.quantity > 0 && stockedLines > 0 ? kit.shippingAmount / stockedLines : 0;
    if (occurrence.shipping !== null) return occurrence.shipping;
    if (occurrence.sourceShippingWithMargin !== null) return occurrence.sourceShippingWithMargin * (1 - state.config.shippingMargin);
    return null;
  }

  // The live rate plus the manufacturer's buffer; the output currency itself never gets a buffer.
  function effectiveFx(currency, margin = 0) {
    const live = fxToOutput(currency);
    if (live === null || !currency || currency === state.config.outputCurrency) return live;
    return live + margin;
  }

  function fxToOutput(currency) {
    if (!currency || currency === state.config.outputCurrency) return 1;
    return state.rates[currency] ?? null;
  }

  function toOutput(amount, currency) {
    if (amount === null || amount === undefined || !Number.isFinite(amount)) return null;
    if (amount === 0) return 0;
    const rate = fxToOutput(currency);
    return rate === null ? null : amount * rate;
  }

  function groupById(id) {
    return id ? state.groups.find((group) => group.id === id) || null : null;
  }

  function effectiveMultiplier(key) {
    const override = state.items[key] || {};
    if (Number.isFinite(override.multiplier)) return override.multiplier;
    const group = groupById(override.group);
    if (Number.isFinite(group?.multiplier)) return group.multiplier;
    return state.config.multiplier;
  }

  function expectedSales(group, quantity) {
    if (!(quantity > 0)) return 0;
    const useGroup = group && Number.isFinite(group.salesValue);
    const mode = useGroup ? group.salesMode : state.sales.defaultMode;
    const value = useGroup ? group.salesValue : state.sales.defaultValue;
    return mode === 'units' ? value : value * quantity;
  }

  function buildConsolidatedItem(occurrences) {
    const occurrenceSummary = Core.summarizeOccurrences(occurrences);
    const sources = occurrenceSummary.sources;
    const ordered = [...occurrences].sort((a, b) => {
      const qtyDiff = (b.quantity ?? -Infinity) - (a.quantity ?? -Infinity);
      if (qtyDiff !== 0) return qtyDiff;
      if (a.price === null && b.price !== null) return 1;
      if (a.price !== null && b.price === null) return -1;
      return a.sheetName.localeCompare(b.sheetName);
    });
    const representative = ordered[0];
    const key = representative.key;
    const override = state.items[key] || {};
    const group = groupById(override.group);
    const sourceMaxQuantity = occurrenceSummary.maxQuantity;
    const quantityOverridden = Number.isFinite(override.quantity);
    const maxQuantity = quantityOverridden ? override.quantity : sourceMaxQuantity;
    const multiplier = effectiveMultiplier(key);
    const fx = effectiveFx(representative.currency, representative.fxMargin);
    const convertedUnit = representative.price !== null && fx !== null ? representative.price * fx : null;
    const convertedTotal = convertedUnit !== null ? convertedUnit * maxQuantity : null;
    const usdToOutput = state.config.outputCurrency === 'USD' ? 1 : state.rates.USD ?? null;
    const amountUsd = convertedTotal !== null && usdToOutput !== null ? convertedTotal / usdToOutput : null;
    const discount = Number.isFinite(representative.discountRate) ? representative.discountRate : state.config.discount;
    const discountedTotal = convertedTotal !== null ? convertedTotal * (1 - discount) : null;
    const sellingTotal = discountedTotal !== null ? discountedTotal * multiplier : null;
    const salesOverridden = Number.isFinite(override.salesPerYear);

    return {
      key,
      part: representative.part,
      description: representative.description,
      // The second article number can differ between kits for the same part, so every one that occurs is listed.
      altParts: [...new Set(occurrences.map((occurrence) => occurrence.altPart).filter(Boolean))],
      common: occurrenceSummary.common,
      sourceMaxQuantity,
      maxQuantity,
      quantityOverridden,
      sources,
      currency: representative.currency,
      unitPrice: representative.price,
      fx,
      fxLive: fxToOutput(representative.currency),
      fxMargin: representative.fxMargin || 0,
      convertedUnit,
      convertedTotal,
      amountUsd,
      discount,
      discountedTotal,
      multiplier,
      multiplierOverridden: multiplier !== state.config.multiplier,
      sellingTotal,
      freightPercent: representative.freightPercent,
      kitFreight: representative.freightPercent !== null && representative.freightPercent !== undefined
        ? (discountedTotal === null ? null : discountedTotal * representative.freightPercent)
        : representative.kitFreight ?? null,
      groupId: group?.id || null,
      groupName: group?.name || '',
      groupColor: group?.color || '',
      salesPerYear: salesOverridden ? override.salesPerYear : expectedSales(group, maxQuantity),
      salesOverridden,
      excluded: override.excluded === true,
      pricingSource: representative.sheetName,
      manufacturer: representative.manufacturer,
      originCountry: representative.originCountry,
      kitDutyRate: representative.dutyRate,
      pricingRow: representative.row + 1,
      occurrenceCount: occurrences.length,
    };
  }

  // Adds consolidated freight, duty, insurance, clearance, VAT and landed cost to each line.
  function applyLandedCosts(items, allocateShipment) {
    const freight = state.freight;
    const margin = state.config.shippingMargin;
    const vatRate = salesVatRate();
    const treatments = items.map((item) => Core.customsTreatment(item.originCountry, state.warehouse.country));
    // Lines share a consolidated shipment only with lines in the same pool: the project-wide pool, or a manufacturer
    // that ships on its own. Clearance and broker fees belong to import declarations only, so they are shared
    // among the imported lines of a pool.
    const shares = new Array(items.length).fill(0);
    const importShares = new Array(items.length).fill(0);
    const shipments = new Array(items.length).fill(0);
    const clearances = new Array(items.length).fill(0);
    const insuranceRates = items.map((item) => makerInsuranceRate(item.manufacturer));
    const pools = new Map();
    items.forEach((item, index) => {
      const key = makerHasOwnShipment(item.manufacturer) ? `maker:${item.manufacturer}` : 'project';
      if (!pools.has(key)) pools.set(key, []);
      pools.get(key).push(index);
    });
    pools.forEach((indexes, key) => {
      const own = key === 'project' ? null : makerSettings(key.slice(6));
      const poolShipment = !allocateShipment ? 0 : own ? toOutput(own.shipmentAmount, own.shipmentCurrency) : toOutput(freight.consolidatedShipment, freight.consolidatedCurrency);
      const poolClearance = !allocateShipment ? 0 : own ? toOutput(own.clearanceAmount, own.clearanceCurrency) : toOutput(freight.clearanceFee, freight.clearanceCurrency);
      const poolItems = indexes.map((index) => items[index]);
      const poolShares = Core.allocationShares(poolItems, freight.allocation);
      const poolImportShares = Core.allocationShares(poolItems.map((item, position) => (treatments[indexes[position]] === 'import' ? item : { maxQuantity: 0, discountedTotal: 0 })), freight.allocation);
      indexes.forEach((index, position) => {
        shares[index] = poolShares[position];
        importShares[index] = poolImportShares[position];
        shipments[index] = poolShipment;
        clearances[index] = poolClearance;
      });
    });

    items.forEach((item, index) => {
      const shipment = shipments[index];
      const clearance = clearances[index];
      const share = allocateShipment ? shares[index] : 0;
      const isImport = treatments[index] === 'import';
      item.routeTreatment = treatments[index];
      const purchase = item.discountedTotal;
      item.allocationShare = share;
      item.consolidatedFreight = shipment === null ? null : shipment * share;
      item.clearance = clearance === null ? null : (isImport && allocateShipment ? clearance * importShares[index] : 0);
      item.insurance = purchase === null ? null : purchase * insuranceRates[index];
      item.missingFreight = item.kitFreight === null || (allocateShipment && (shipment === null || clearance === null));
      const costComponents = [purchase, item.kitFreight, item.consolidatedFreight, item.insurance, item.clearance];
      const costsComplete = costComponents.every((value) => value !== null && Number.isFinite(value));
      const customsValue = costsComplete ? purchase + item.kitFreight + item.consolidatedFreight + item.insurance : null;
      item.dutyRate = isImport ? (Number.isFinite(item.kitDutyRate) ? item.kitDutyRate : freight.dutyRate) : 0;
      item.duty = customsValue === null ? null : customsValue * item.dutyRate;
      item.importCosts = costsComplete ? item.insurance + item.duty + item.clearance : null;
      item.freightAndImport = costsComplete ? item.kitFreight + item.consolidatedFreight + item.importCosts : null;
      item.shippingMargin = margin;
      item.freightWithMargin = item.freightAndImport === null ? null : item.freightAndImport / (1 - margin);
      item.lineTotal = item.sellingTotal === null || item.freightWithMargin === null ? null : item.sellingTotal + item.freightWithMargin;
      item.vat = item.lineTotal === null ? null : item.lineTotal * vatRate;
      item.lineTotalInclVat = item.lineTotal === null ? null : item.lineTotal + item.vat;
      item.landedCost = purchase === null || item.freightAndImport === null ? null : purchase + item.freightAndImport;
      item.importVat = customsValue === null || item.duty === null ? null : (isImport ? (customsValue + item.duty) * vatRate : 0);
      item.unitSalesPrice = item.lineTotal !== null && item.maxQuantity > 0 ? item.lineTotal / item.maxQuantity : null;
      item.unitLandedCost = item.landedCost !== null && item.maxQuantity > 0 ? item.landedCost / item.maxQuantity : null;
      // Legacy names used by the warehouse sheet and older exports.
      item.shipping = item.kitFreight;
      item.shippingWithMargin = item.freightWithMargin;
    });
  }

  // Each kit on its own, as if it were ordered and shipped separately.
  function buildKitSummaries(kitOccurrences) {
    const summaries = [];
    kitOccurrences.forEach((occurrences, sheetName) => {
      const kit = kitSettings(sheetName);
      let list = 0;
      let purchase = 0;
      let sales = 0;
      let freight = 0;
      let missing = 0;
      occurrences.forEach((occurrence) => {
        const fx = effectiveFx(occurrence.currency, occurrence.fxMargin);
        if (occurrence.price === null || fx === null) {
          missing += 1;
        } else {
          const lineList = occurrence.price * fx * occurrence.quantity;
          const linePurchase = lineList * (1 - kitDiscount(kit));
          list += lineList;
          purchase += linePurchase;
          sales += linePurchase * effectiveMultiplier(occurrence.key);
        }
        freight += occurrence.kitFreight ?? 0;
      });
      summaries.push({
        sheetName,
        lines: occurrences.length,
        stockedLines: occurrences.filter((occurrence) => occurrence.quantity > 0).length,
        units: occurrences.reduce((sum, occurrence) => sum + Math.max(0, occurrence.quantity), 0),
        list,
        purchase,
        freight,
        salesInclFreight: sales + freight / (1 - state.config.shippingMargin),
        missing,
        mode: kitSettings(sheetName).shippingMode,
        manufacturer: kit.manufacturer,
        originCountry: kit.originCountry,
        currency: kitCurrency(kit),
        discountRate: kitDiscount(kit),
      });
    });
    return summaries;
  }

  function readCalculatedValue(sheetName, row, col) {
    if (col === null || col === undefined) return null;
    try {
      const value = state.hf.getCellValue({ sheet: state.sheetIds[sheetName], row, col });
      if (isFormulaError(value)) return null;
      return value;
    } catch {
      return state.matrices[sheetName]?.[row]?.[col] ?? null;
    }
  }

  /* ---------- View refresh ---------- */

  function toggleSetupGuide(force) {
    if (!state.workbook) return;
    const show = typeof force === 'boolean' ? force : state.guideDismissed;
    state.guideDismissed = !show;
    renderSetupGuide();
  }

  function setupGuideSteps() {
    const enabledMappings = Object.values(state.mappings).filter((mapping) => mapping.enabled);
    const mapped = enabledMappings.filter((mapping) => (
      mapping.part !== null
      && mapping.quantity !== null
      && (mapping.price !== null || mapping.discountedTotal !== null || mapping.sellingTotal !== null)
    )).length;
    const mappingIssues = Math.max(0, enabledMappings.length - mapped);
    const missingPricesOrRates = state.consolidated.filter((item) => item.convertedTotal === null).length;
    const missingFreight = state.consolidated.filter((item) => item.missingFreight).length;
    const warehouseConfigured = Object.values(state.warehouse.rates).some((value) => Number(value) > 0);
    const warehouseFxReady = state.config.outputCurrency === 'NOK' || Number.isFinite(state.rates.NOK);
    const profitability = calculateProfitability();
    return [
      {
        complete: enabledMappings.length > 0 && mappingIssues === 0,
        title: 'Confirm the parts sheets',
        detail: !enabledMappings.length ? 'No parts sheets are enabled.' : mappingIssues ? `${mappingIssues} enabled sheet${mappingIssues === 1 ? '' : 's'} need a part, quantity and price column.` : `${mapped} parts sheet${mapped === 1 ? '' : 's'} mapped and included.`,
        action: 'mapping',
        actionLabel: 'Review sheets',
      },
      {
        complete: state.consolidated.length > 0 && missingPricesOrRates === 0,
        title: 'Check prices and currencies',
        detail: missingPricesOrRates ? `${missingPricesOrRates} part${missingPricesOrRates === 1 ? '' : 's'} have a missing price or currency conversion.` : `${state.consolidated.length} consolidated parts have usable prices.`,
        action: 'rates',
        actionLabel: 'Check rates',
      },
      {
        complete: state.consolidated.length > 0 && missingFreight === 0,
        title: 'Confirm inbound freight',
        detail: missingFreight ? `${missingFreight} part${missingFreight === 1 ? '' : 's'} still need freight data.` : 'Kit freight and shared shipment inputs are calculable.',
        action: 'freight',
        actionLabel: 'Review freight',
      },
      {
        complete: warehouseConfigured && warehouseFxReady,
        title: 'Review warehouse assumptions',
        detail: !warehouseConfigured ? 'Warehouse quote rates have not been entered or imported.' : !warehouseFxReady ? `The NOK to ${state.config.outputCurrency} rate is missing.` : 'Warehouse quote rates and NOK conversion are ready.',
        action: 'warehouse',
        actionLabel: 'Open warehouse',
      },
      {
        complete: profitability.complete,
        title: 'Review the business result',
        detail: profitability.complete ? 'Profit, margin, ROI and payback can be calculated.' : `The estimate is incomplete: ${profitability.missingReasons.join(' and ')}.`,
        action: 'dashboard',
        actionLabel: 'Open dashboard',
      },
      {
        complete: !state.projectDirty,
        title: 'Save a shareable copy',
        detail: state.projectDirty ? 'Your latest changes have not been exported. Use an encrypted project when access needs a password.' : state.projectSaveLabel || 'The current workbook has been imported.',
        action: 'save',
        actionLabel: 'Save options',
      },
    ];
  }

  function renderSetupGuide() {
    const available = Boolean(state.workbook);
    dom.setupGuide.classList.toggle('hidden', !available || state.guideDismissed);
    dom.guideToggle.disabled = !available;
    dom.guideToggle.setAttribute('aria-expanded', String(available && !state.guideDismissed));
    dom.guideToggle.textContent = available && !state.guideDismissed ? 'Hide guide' : 'Setup guide';
    if (!available) return;
    const steps = setupGuideSteps();
    const complete = steps.filter((step) => step.complete).length;
    const percent = Math.round(complete / steps.length * 100);
    if (state.guideDismissed) dom.guideToggle.textContent = `Setup guide ${complete}/${steps.length}`;
    dom.setupProgress.setAttribute('aria-valuenow', String(percent));
    dom.setupProgressBar.style.width = `${percent}%`;
    dom.setupGuideSummary.textContent = complete === steps.length
      ? 'Everything needed for a complete estimate is ready. Review the result, then export the format you need.'
      : `${complete} of ${steps.length} checks are complete. Work from top to bottom; advanced settings can stay unchanged.`;
    const rows = steps.map((step, index) => {
      const row = document.createElement('li');
      row.className = `setup-step ${step.complete ? 'complete' : 'warning'}`;
      const icon = document.createElement('span');
      icon.className = 'setup-step-icon';
      icon.textContent = step.complete ? '✓' : String(index + 1);
      const copy = document.createElement('span');
      copy.className = 'setup-step-copy';
      const title = document.createElement('strong');
      title.textContent = step.title;
      const detail = document.createElement('span');
      detail.textContent = step.detail;
      copy.append(title, detail);
      const action = document.createElement('button');
      action.type = 'button';
      action.className = `button small ${step.complete ? 'ghost' : ''}`;
      action.dataset.guideAction = step.action;
      action.textContent = step.actionLabel;
      row.append(icon, copy, action);
      return row;
    });
    dom.setupChecklist.replaceChildren(...rows);
  }

  function handleSetupGuideAction(event) {
    const button = event.target.closest('[data-guide-action]');
    if (!button) return;
    const action = button.dataset.guideAction;
    if (action === 'warehouse') showView(VIEW.warehouse);
    else if (action === 'dashboard') showView(VIEW.dashboard);
    else {
      const target = {
        mapping: document.getElementById('mapping-panel'),
        rates: document.getElementById('workbook-panel'),
        freight: document.getElementById('freight-panel'),
        save: document.getElementById('secure-panel'),
      }[action];
      target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      target?.querySelector('button, input, select')?.focus({ preventScroll: true });
    }
    toggleSetupGuide(false);
  }

  function refreshViews() {
    if (!state.workbook) return;
    renderSaveState();
    renderSetupGuide();
    renderKitBars();
    renderMakers();
    renderFreightSummary();
    updateGroupCounts();
    populateOriginSelects();
    populateCustomerSelects();
    renderVatControls();
    updateKitRouteSummaries();
    renderCustomsGuides();
    if (state.activeView === VIEW.consolidated) renderConsolidated();
    else if (state.activeView === VIEW.dashboard) renderDashboard();
    else if (state.activeView === VIEW.warehouse) renderWarehouse();
    else if (state.activeView === VIEW.wire) {
      renderFlow();
      renderWireView();
    }
    syncBoundInputs();
  }

  function renderFreightSummary() {
    const output = state.config.outputCurrency;
    const kitFreight = state.consolidated.reduce((sum, item) => sum + (item.kitFreight ?? 0), 0);
    const shipment = state.consolidated.reduce((sum, item) => sum + (item.consolidatedFreight ?? 0), 0);
    const importCosts = state.consolidated.reduce((sum, item) => sum + (item.importCosts ?? 0), 0);
    const missingRate = state.consolidated.some((item) => item.missingFreight);
    dom.freightSummary.textContent = `Consolidated order: kit freight ${formatMoney(kitFreight, output)} + shipment ${formatMoney(shipment, output)} + duty, insurance & fees ${formatMoney(importCosts, output)}.${missingRate ? ' Some currency rates are still missing.' : ''}`;
  }

  function renderKitBars() {
    [dom.kitBar, dom.dashboardKitBar].forEach((bar) => {
      const label = document.createElement('span');
      label.className = 'kit-bar-label';
      label.textContent = 'Kits';
      const makers = [...new Set(state.workbook.SheetNames.filter(isSourceKit).map((name) => kitSettings(name).manufacturer || name))];
      const quick = makers.length > 1 ? [['', 'All kits'], ...makers.map((maker) => [maker, `Only ${maker}`])].map(([maker, text]) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'kit-quick';
        button.dataset.kitOnly = maker;
        button.textContent = text;
        return button;
      }) : [];
      const chips = state.workbook.SheetNames.filter(isSourceKit).map((sheetName) => {
        const enabled = state.mappings[sheetName].enabled;
        const summary = state.kitSummaries.find((item) => item.sheetName === sheetName);
        const chip = document.createElement(enabled ? 'span' : 'button');
        chip.className = `kit-chip${enabled ? '' : ' removed'}`;
        if (enabled) {
          const name = document.createElement('strong');
          name.textContent = sheetName;
          name.translate = false;
          const meta = document.createElement('small');
          meta.textContent = summary ? `${summary.lines} lines` : '';
          const remove = document.createElement('button');
          remove.type = 'button';
          remove.className = 'kit-chip-remove';
          remove.dataset.kitToggle = sheetName;
          remove.dataset.kitEnable = 'false';
          remove.setAttribute('aria-label', `Remove ${sheetName} from the consolidation`);
          remove.title = `Remove ${sheetName} and recalculate`;
          remove.textContent = '×';
          chip.append(name, meta, remove);
        } else {
          chip.type = 'button';
          chip.dataset.kitToggle = sheetName;
          chip.dataset.kitEnable = 'true';
          chip.title = `Add ${sheetName} back and recalculate`;
          chip.textContent = `+ ${sheetName}`;
        }
        return chip;
      });
      if (!chips.length) {
        const empty = document.createElement('span');
        empty.className = 'muted';
        empty.textContent = 'No kit sheets detected.';
        chips.push(empty);
      }
      bar.replaceChildren(label, ...quick, ...chips);
    });
    renderKitBreakdown();
  }

  function renderKitBreakdown() {
    const output = state.config.outputCurrency;
    const table = document.createElement('table');
    table.className = 'mini-table';
    const head = table.createTHead().insertRow();
    ['Kit', 'Lines', 'Units', 'Purchase after discount', 'Kit freight', 'Freight mode', `Sales incl. freight (${output})`].forEach((text) => {
      const th = document.createElement('th');
      th.textContent = text;
      head.append(th);
    });
    const body = table.createTBody();
    state.kitSummaries.forEach((summary) => {
      const row = body.insertRow();
      [summary.sheetName, `${summary.lines} (${summary.stockedLines} stocked)`, formatNumber(summary.units, 0), formatMoney(summary.purchase, output), formatMoney(summary.freight, output), KIT_SHIPPING_MODES[summary.mode], formatMoney(summary.salesInclFreight, output)]
        .forEach((value, index) => {
          const cell = row.insertCell();
          cell.textContent = value;
          if (index > 0 && index !== 5) cell.className = 'number';
        });
    });
    const separateFreight = state.kitSummaries.reduce((sum, summary) => sum + summary.freight, 0);
    const consolidatedFreight = state.consolidated.reduce((sum, item) => sum + (item.kitFreight ?? 0) + (item.consolidatedFreight ?? 0), 0);
    const note = document.createElement('p');
    note.className = 'fine-print';
    note.textContent = `Kits shipped separately: ${formatMoney(separateFreight, output)} kit freight. Consolidated order: ${formatMoney(consolidatedFreight, output)} (kit freight on the consolidated lines plus the consolidated shipment).`;
    dom.kitBreakdownTable.replaceChildren(table, note);
  }

  /* ---------- Consolidated table ---------- */

  // Every column carries a stable `key` (used for sorting, the chooser and the hidden list) and an
  // optional `sortValue` giving the raw number or text to order by. `value` stays the display string,
  // because most columns are formatted for reading rather than for comparing.
  function consolidatedColumns() {
    const o = state.config.outputCurrency;
    const money = (field) => (item) => formatNullableNumber(item[field], 2);
    return [
      { key: 'sources', label: 'Source sheets', data: true, sortValue: (item) => item.sources.join(', '), value: (item) => item.sources.join(', ') },
      { key: 'part', label: 'Part', data: true, sortValue: (item) => item.part, value: (item) => item.part, className: 'part-cell' },
      { key: 'description', label: 'Description', data: true, sortValue: (item) => item.description, value: (item) => item.description },
      { key: 'altPart', label: 'Alt. part no.', data: true, optional: true, sortValue: (item) => item.altParts.join(', '), value: (item) => item.altParts.join(', ') },
      { key: 'quantity', label: 'Quantity', number: true, sortValue: (item) => item.maxQuantity, value: (item) => formatNumber(item.maxQuantity, 2), overridden: (item) => item.quantityOverridden, note: (item) => (item.quantityOverridden ? `Stock quantity override. Source maximum: ${formatNumber(item.sourceMaxQuantity, 2)}` : '') },
      { key: 'convertedTotal', label: `Amount in ${o}`, number: true, sortValue: (item) => item.convertedTotal, value: money('convertedTotal'), required: true },
      { key: 'manufacturer', label: 'Manufacturer', data: true, sortValue: (item) => item.manufacturer || item.pricingSource, value: (item) => item.manufacturer || item.pricingSource },
      { key: 'originCountry', label: 'Country of origin', sortValue: (item) => originName(item.originCountry), value: (item) => originName(item.originCountry) },
      { key: 'match', label: 'Match', sortValue: (item) => (item.common ? 1 : 0), value: (item) => (item.common ? 'Common' : 'Unique') },
      { key: 'salesGroup', label: 'Sales group', group: true, sortValue: (item) => item.groupName || '' },
      { key: 'salesPerYear', label: 'Expected sales / yr', number: true, sortValue: (item) => item.salesPerYear, value: (item) => formatNumber(item.salesPerYear, 2), overridden: (item) => item.salesOverridden, note: (item) => (item.salesOverridden ? 'Item sales override' : item.groupName ? `From group ${item.groupName}` : 'Default sales assumption') },
      { key: 'currency', label: 'Currency', sortValue: (item) => item.currency, value: (item) => item.currency },
      { key: 'unitPrice', label: 'Unit price', number: true, sortValue: (item) => item.unitPrice, value: money('unitPrice'), required: true },
      { key: 'fx', label: `FX to ${o}`, number: true, sortValue: (item) => item.fx, value: (item) => formatNullableNumber(item.fx, 6), required: true },
      { key: 'convertedUnit', label: `Unit in ${o}`, number: true, sortValue: (item) => item.convertedUnit, value: money('convertedUnit'), required: true },
      { key: 'discount', label: 'Discount', number: true, sortValue: (item) => item.discount, value: (item) => formatPercent(item.discount) },
      { key: 'discountedTotal', label: `Purchase after discount in ${o}`, number: true, sortValue: (item) => item.discountedTotal, value: money('discountedTotal'), required: true },
      { key: 'multiplier', label: 'Price multiplier', number: true, sortValue: (item) => item.multiplier, value: (item) => formatNumber(item.multiplier, 2), overridden: (item) => item.multiplierOverridden },
      { key: 'sellingTotal', label: `Sales price in ${o}`, number: true, sortValue: (item) => item.sellingTotal, value: money('sellingTotal'), required: true },
      { key: 'kitFreight', label: `Kit freight in ${o}`, number: true, sortValue: (item) => item.kitFreight, value: money('kitFreight'), required: true },
      { key: 'consolidatedFreight', label: `Consolidated shipment in ${o}`, number: true, sortValue: (item) => item.consolidatedFreight, value: money('consolidatedFreight') },
      { key: 'importCosts', label: `Duty, insurance & fees in ${o}`, number: true, sortValue: (item) => item.importCosts, value: (item) => formatNumber(item.importCosts, 2) },
      { key: 'shippingMargin', label: 'Freight margin', number: true, sortValue: (item) => item.shippingMargin, value: (item) => formatPercent(item.shippingMargin) },
      { key: 'freightWithMargin', label: `Freight & import incl. margin in ${o}`, number: true, sortValue: (item) => item.freightWithMargin, value: (item) => formatNumber(item.freightWithMargin, 2) },
      { key: 'lineTotal', label: salesVatRate() > 0 ? `Line total excl. VAT in ${o}` : `Line total in ${o}`, number: true, sortValue: (item) => item.lineTotal, value: money('lineTotal'), required: true, strong: true },
      { key: 'vat', label: `VAT in ${o}`, number: true, vat: true, sortValue: (item) => item.vat, value: money('vat') },
      { key: 'lineTotalInclVat', label: `Line total incl. VAT in ${o}`, number: true, vat: true, sortValue: (item) => item.lineTotalInclVat, value: money('lineTotalInclVat'), required: true },
      { key: 'landedCost', label: `Landed cost in ${o}`, number: true, sortValue: (item) => item.landedCost, value: money('landedCost'), required: true },
      { key: 'pricingSource', label: 'Pricing source', sortValue: (item) => `${item.pricingSource} row ${item.pricingRow}`, value: (item) => `${item.pricingSource} row ${item.pricingRow}` },
    ];
  }

  // Columns the user has not hidden, minus the VAT columns when VAT is off.
  function visibleConsolidatedColumns() {
    const hidden = new Set(state.hiddenColumns);
    const hasAltParts = state.consolidated.some((item) => item.altParts.length);
    return consolidatedColumns().filter((column) => !hidden.has(column.key) && (!column.vat || salesVatRate() > 0) && (!column.optional || hasAltParts));
  }

  // Builds the show/hide list in the Columns menu. Hidden columns are listed too, ticked off, so a
  // column that was switched off by mistake can always be found and switched back on.
  function renderColumnChooser(columns) {
    if (!columns.length) {
      dom.columnList.replaceChildren();
      return;
    }
    dom.columnList.replaceChildren(...columns.map((column) => {
      const label = document.createElement('label');
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.checked = !state.hiddenColumns.includes(column.key);
      box.dataset.columnKey = column.key;
      box.setAttribute('aria-label', column.label);
      const text = document.createElement('span');
      text.textContent = column.label;
      label.append(box, text);
      if (!box.checked) label.classList.add('is-hidden-column');
      return label;
    }));
  }

  function toggleColumn(key, visible) {
    const next = visible ? state.hiddenColumns.filter((item) => item !== key) : [...new Set([...state.hiddenColumns, key])];
    state.hiddenColumns = next;
    saveUiPreferences();
    if (state.activeView === VIEW.consolidated) renderConsolidated();
  }


  const makerOf = (item) => item.manufacturer || item.pricingSource;

  function itemMatchesFilter(item, filter, term) {
    if (term) {
      const haystack = `${item.part} ${item.altParts.join(' ')} ${item.description} ${item.sources.join(' ')} ${item.groupName}`.toLowerCase();
      if (!haystack.includes(term)) return false;
    }
    if (filter === 'common') return item.common;
    if (filter === 'unique') return !item.common;
    if (filter === 'ungrouped') return !item.groupId;
    if (filter === 'selected') return state.selectedKeys.has(item.key);
    if (filter === 'overridden') return item.quantityOverridden || item.salesOverridden || item.multiplierOverridden;
    if (filter === 'missing') return Core.hasMissingInputs(item);
    if (filter.startsWith('group:')) return item.groupId === filter.slice(6);
    if (filter.startsWith('maker:')) return makerOf(item) === filter.slice(6);
    if (filter.startsWith('makercommon:')) return item.common && makerOf(item) === filter.slice(12);
    if (filter.startsWith('makerunique:')) return !item.common && makerOf(item) === filter.slice(12);
    return true;
  }

  function renderConsolidatedFilter() {
    const makers = [...new Set(state.consolidated.map(makerOf).filter(Boolean))];
    const options = [
      ['all', 'All parts'], ['common', 'Common parts'], ['unique', 'Unique parts'], ['selected', 'Selected'],
      ['overridden', 'With overrides'], ['ungrouped', 'No sales group'], ['missing', 'Missing inputs'],
      ...state.groups.map((group) => [`group:${group.id}`, `Group: ${group.name}`]),
      // With several manufacturers in one warehouse, each one gets its own common and unique lists.
      ...(makers.length > 1 ? makers.flatMap((maker) => [[`maker:${maker}`, `Manufacturer: ${maker}`], [`makercommon:${maker}`, `${maker}: common parts`], [`makerunique:${maker}`, `${maker}: unique parts`]]) : []),
      ['excluded', 'Excluded parts'],
    ];
    if (!options.some(([value]) => value === state.consolidatedFilter)) state.consolidatedFilter = 'all';
    dom.consolidatedFilter.replaceChildren(...options.map(([value, label]) => new Option(label, value, false, value === state.consolidatedFilter)));
    dom.consolidatedFilter.value = state.consolidatedFilter;
  }

  function renderConsolidated() {
    renderConsolidatedFilter();
    const output = state.config.outputCurrency;
    const items = state.consolidated;
    const common = items.filter((item) => item.common);
    const unique = items.filter((item) => !item.common);
    const missingPrices = items.filter((item) => item.unitPrice === null || item.fx === null).length;
    const missingFreight = items.filter((item) => item.missingFreight).length;
    const sum = (field) => items.reduce((total, item) => total + (item[field] ?? 0), 0);

    dom.consolidatedSummary.textContent = `${items.length} parts from ${enabledSheetCount()} source sheets${state.excludedItems.length ? ` · ${state.excludedItems.length} excluded` : ''}. Common parts are highlighted yellow.`;
    dom.summaryCards.replaceChildren(
      summaryCard('Common parts', common.length),
      summaryCard('Unique parts', unique.length),
      summaryCard(salesVatRate() > 0 ? `Total excl. VAT (${output})` : `Total (${output})`, formatNumber(sum('lineTotal'), 2)),
      ...(salesVatRate() > 0 ? [summaryCard(`Total incl. VAT (${output})`, formatNumber(sum('lineTotalInclVat'), 2))] : []),
      summaryCard(`Landed cost (${output})`, formatNumber(sum('landedCost'), 2)),
      // Clicking the count filters the table down to the rows that still need something filled in,
      // the same shortcut pattern the setup guide already uses for its own checks.
      actionCard('Missing inputs', missingPrices + missingFreight, () => {
        state.consolidatedFilter = 'missing';
        renderConsolidated();
      }, {
        attention: missingPrices + missingFreight > 0,
        hint: missingPrices + missingFreight > 0
          ? 'Show only the parts with a missing price or freight'
          : 'Every part has a price and freight',
      }),
    );

    const term = state.consolidatedSearch.trim().toLowerCase();
    const filter = state.consolidatedFilter;
    const showExcluded = filter === 'excluded';
    // The chosen sort column orders rows inside each block; the Common/Unique grouping is kept so the
    // table still reads top-down, and an unsorted column leaves the default common-then-part order alone.
    const sortColumns = visibleConsolidatedColumns();
    const sortColumn = sortColumns.find((column) => column.key === state.consolidatedSort?.key);
    const sorter = (list) => (sortColumn
      ? Core.sortItems(list, state.consolidatedSort, { keyFor: sortColumn.sortValue || sortColumn.value, tieBreaker: (item) => item.part })
      : list);
    const visibleCommon = showExcluded ? [] : sorter(common.filter((item) => itemMatchesFilter(item, filter, term)));
    const visibleUnique = showExcluded ? [] : sorter(unique.filter((item) => itemMatchesFilter(item, filter, term)));
    const visibleExcluded = sorter(state.excludedItems.filter((item) => itemMatchesFilter(item, showExcluded ? 'all' : filter, term) && (showExcluded || filter === 'all' || filter === 'selected')));
    state.visibleKeys = [...visibleCommon, ...visibleUnique, ...visibleExcluded].map((item) => item.key);
    const visibleCount = visibleCommon.length + visibleUnique.length;
    dom.consolidatedCount.textContent = showExcluded
      ? `${visibleExcluded.length} excluded`
      : `${visibleCount} of ${items.length} shown${state.selectedKeys.size ? ` · ${state.selectedKeys.size} selected` : ''}`;

    const columns = visibleConsolidatedColumns();
    renderColumnChooser(consolidatedColumns());
    const table = document.createElement('table');
    table.className = 'parts-table';
    const headRow = table.createTHead().insertRow();
    const selectHead = document.createElement('th');
    selectHead.className = 'select-cell';
    const selectAll = document.createElement('input');
    selectAll.type = 'checkbox';
    selectAll.dataset.selectAll = 'true';
    selectAll.setAttribute('aria-label', 'Select all visible parts');
    const visibleSelected = state.visibleKeys.filter((key) => state.selectedKeys.has(key)).length;
    selectAll.checked = state.visibleKeys.length > 0 && visibleSelected === state.visibleKeys.length;
    selectAll.indeterminate = visibleSelected > 0 && visibleSelected < state.visibleKeys.length;
    selectHead.append(selectAll);
    headRow.append(selectHead);
    columns.forEach((column) => {
      const th = document.createElement('th');
      th.className = column.number ? 'number sortable' : 'sortable';
      th.scope = 'col';
      th.dataset.sortKey = column.key;
      const text = document.createElement('span');
      text.textContent = column.label;
      const caret = document.createElement('span');
      caret.className = 'sort-caret';
      caret.setAttribute('aria-hidden', 'true');
      th.append(text, caret);
      const sorted = state.consolidatedSort?.key === column.key;
      th.setAttribute('aria-sort', sorted ? (state.consolidatedSort.direction === 'asc' ? 'ascending' : 'descending') : 'none');
      th.setAttribute('aria-label', `${column.label}: sort ${sorted && state.consolidatedSort.direction === 'asc' ? 'descending' : 'ascending'}`);
      th.tabIndex = 0;
      headRow.append(th);
    });

    const tbody = table.createTBody();
    const columnCount = columns.length + 1;
    if (!items.length && !state.excludedItems.length) {
      const row = tbody.insertRow();
      row.className = 'empty-row';
      const cell = row.insertCell();
      cell.colSpan = columnCount;
      cell.textContent = 'No parts are available yet. Enable kit sheets and map the Part / SKU and Quantity columns.';
    } else if (!state.visibleKeys.length) {
      const row = tbody.insertRow();
      row.className = 'empty-row';
      const cell = row.insertCell();
      cell.colSpan = columnCount;
      cell.textContent = 'No parts match the current search or filter.';
    } else {
      appendPartGroup(tbody, 'Common parts', visibleCommon, columns);
      appendPartGroup(tbody, 'Unique parts', visibleUnique, columns);
      appendPartGroup(tbody, 'Excluded parts · not in totals', visibleExcluded, columns, true);
    }
    dom.consolidatedTable.replaceChildren(table);
  }

  function appendPartGroup(tbody, label, items, columns, excluded = false) {
    if (!items.length) return;
    const groupRow = tbody.insertRow();
    groupRow.className = 'group-row';
    const groupSelectCell = groupRow.insertCell();
    groupSelectCell.className = 'select-cell';
    groupSelectCell.setAttribute('aria-hidden', 'true');
    const groupCell = groupRow.insertCell();
    groupCell.colSpan = columns.length;
    groupCell.textContent = `${label} — ${items.length}`;

    items.forEach((item) => {
      const row = tbody.insertRow();
      row.className = [item.common ? 'common' : '', excluded ? 'excluded' : '', state.selectedKeys.has(item.key) ? 'selected' : ''].filter(Boolean).join(' ');
      // The whole row is a selection target, so parts can be picked with a click or the keyboard
      // instead of only by finding the small checkbox.
      row.dataset.rowKey = item.key;
      row.tabIndex = state.focusedRowKey === item.key ? 0 : -1;
      row.setAttribute('aria-selected', String(state.selectedKeys.has(item.key)));
      const selectCell = row.insertCell();
      selectCell.className = 'select-cell';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.dataset.selectKey = item.key;
      checkbox.checked = state.selectedKeys.has(item.key);
      checkbox.setAttribute('aria-label', `Select ${item.part}`);
      selectCell.append(checkbox);

      columns.forEach((column) => {
        const cell = row.insertCell();
        if (column.group) {
          if (item.groupId) {
            const chip = document.createElement('span');
            chip.className = 'group-chip';
            chip.style.setProperty('--group-color', item.groupColor);
            chip.textContent = item.groupName;
            chip.translate = false;
            cell.append(chip);
          } else {
            cell.textContent = '—';
            cell.classList.add('muted-cell');
          }
          return;
        }
        const value = column.value(item);
        cell.textContent = value;
        if (column.data) cell.translate = false;
        if (column.className) cell.classList.add(column.className);
        if (column.number) cell.classList.add('number');
        if (column.strong) cell.classList.add('strong-cell');
        if (column.overridden?.(item)) cell.classList.add('overridden');
        const note = column.note?.(item);
        if (note) cell.title = note;
        if (column.required && value === '—') cell.classList.add('warning-text');
      });
    });
  }

  function handleConsolidatedTableClick(event) {
    const header = event.target.closest('th[data-sort-key]');
    if (header) {
      cycleConsolidatedSort(header.dataset.sortKey);
      return;
    }
    const checkbox = event.target.closest('input[type="checkbox"]');
    if (checkbox && checkbox.dataset.selectAll) {
      const keys = state.visibleKeys || [];
      if (checkbox.checked) keys.forEach((key) => state.selectedKeys.add(key));
      else keys.forEach((key) => state.selectedKeys.delete(key));
      state.selectionAnchor = null;
      renderConsolidated();
      renderSelectionBar(dom.consolidatedActions, [...state.selectedKeys], 'consolidated');
      return;
    }
    const row = event.target.closest('tr[data-row-key]');
    if (!row) return;
    const key = row.dataset.rowKey;
    // A click on the checkbox itself already handled the selection, so only row clicks get here.
    if (event.target.closest('input[type="checkbox"]')) return;
    selectConsolidatedRow(key, { range: event.shiftKey, additive: event.ctrlKey || event.metaKey, focus: true });
  }

  // Clicking a header sorts by it: first click ascending, second descending, third back to the
  // default common-then-part order.
  function cycleConsolidatedSort(key) {
    const current = state.consolidatedSort;
    if (!current || current.key !== key) state.consolidatedSort = { key, direction: 'asc' };
    else if (current.direction === 'asc') state.consolidatedSort = { key, direction: 'desc' };
    else state.consolidatedSort = null;
    saveUiPreferences();
    renderConsolidated();
    renderSelectionBar(dom.consolidatedActions, [...state.selectedKeys], 'consolidated');
  }

  function selectConsolidatedRow(key, options = {}) {
    const keys = state.visibleKeys || [];
    const alreadySelected = state.selectedKeys.has(key);
    if (options.range && state.selectionAnchor && keys.includes(state.selectionAnchor)) {
      const [from, to] = [keys.indexOf(state.selectionAnchor), keys.indexOf(key)].sort((a, b) => a - b);
      keys.slice(from, to + 1).forEach((rangeKey) => state.selectedKeys.add(rangeKey));
    } else if (options.additive) {
      if (alreadySelected) state.selectedKeys.delete(key);
      else state.selectedKeys.add(key);
    } else if (alreadySelected) {
      state.selectedKeys.clear();
    } else {
      state.selectedKeys.add(key);
    }
    state.selectionAnchor = key;
    state.focusedRowKey = key;
    renderConsolidated();
    renderSelectionBar(dom.consolidatedActions, [...state.selectedKeys], 'consolidated');
    if (options.focus) dom.consolidatedTable.querySelector(`tr[data-row-key="${CSS.escape(key)}"]`)?.focus();
  }

  // Keyboard selection inside the table: Space toggles, arrows move, Ctrl+A selects what is visible.
  function handleConsolidatedTableKeydown(event) {
    const header = event.target.closest('th[data-sort-key]');
    if (header && (event.key === 'Enter' || event.key === ' ')) {
      event.preventDefault();
      cycleConsolidatedSort(header.dataset.sortKey);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'a') {
      const keys = state.visibleKeys || [];
      if (!keys.length) return;
      event.preventDefault();
      keys.forEach((key) => state.selectedKeys.add(key));
      renderConsolidated();
      renderSelectionBar(dom.consolidatedActions, [...state.selectedKeys], 'consolidated');
      return;
    }
    const row = event.target.closest('tr[data-row-key]');
    if (!row) return;
    const key = row.dataset.rowKey;
    if (event.key === ' ' || event.key === 'Enter') {
      event.preventDefault();
      selectConsolidatedRow(key, { additive: true });
      return;
    }
    const keys = state.visibleKeys || [];
    const index = keys.indexOf(key);
    let nextIndex = null;
    if (event.key === 'ArrowDown') nextIndex = Math.min(keys.length - 1, index + 1);
    if (event.key === 'ArrowUp') nextIndex = Math.max(0, index - 1);
    if (event.key === 'Home') nextIndex = 0;
    if (event.key === 'End') nextIndex = keys.length - 1;
    if (nextIndex === null) return;
    event.preventDefault();
    const nextKey = keys[nextIndex];
    state.focusedRowKey = nextKey;
    if (event.shiftKey && state.selectionAnchor) selectConsolidatedRow(nextKey, { range: true });
    else {
      dom.consolidatedTable.querySelector(`tr[data-row-key="${CSS.escape(nextKey)}"]`)?.focus();
    }
  }

  /* ---------- Item selection & bulk actions ---------- */

  const SELECTION_ACTIONS = [
    ['group', 'Assign to sales group'],
    ['sales', 'Set expected sales / yr'],
    ['quantity', 'Set stock quantity'],
    ['multiplier', 'Set price multiplier'],
    ['exclude', 'Exclude from consolidation'],
    ['include', 'Include in consolidation'],
    ['clear', 'Clear quantity, sales & multiplier overrides'],
  ];

  function renderAllSelectionBars() {
    renderSelectionBar(dom.consolidatedActions, [...state.selectedKeys], 'consolidated');
    renderSelectionBar(dom.sheetActions, sheetSelectionKeys(), 'sheet');
  }

  function renderSelectionBarCount() {
    const count = dom.consolidatedActions.querySelector('.selection-count');
    if (count) count.textContent = `${state.selectedKeys.size} part${state.selectedKeys.size === 1 ? '' : 's'} selected`;
  }

  function renderSelectionBar(container, keys, context) {
    container.dataset.context = context;
    const rowCount = context === 'sheet' ? state.sheetSelection.rows.size : keys.length;
    if (!rowCount) {
      container.classList.add('hidden');
      container.replaceChildren();
      return;
    }
    container.classList.remove('hidden');
    const count = document.createElement('strong');
    count.className = 'selection-count';
    count.textContent = context === 'sheet'
      ? `${rowCount} row${rowCount === 1 ? '' : 's'} · ${keys.length} part${keys.length === 1 ? '' : 's'}`
      : `${keys.length} part${keys.length === 1 ? '' : 's'} selected`;
    const clear = document.createElement('button');
    clear.type = 'button';
    clear.className = 'button small ghost';
    clear.dataset.selectionClear = 'true';
    clear.textContent = 'Clear selection';
    if (!keys.length) {
      const note = document.createElement('span');
      note.className = 'selection-note';
      note.textContent = 'The selected rows contain no mapped part numbers.';
      container.replaceChildren(count, note, clear);
      return;
    }
    const action = document.createElement('select');
    action.className = 'action-select';
    action.setAttribute('aria-label', 'Action for the selected parts');
    SELECTION_ACTIONS.forEach(([value, label]) => action.append(new Option(label, value, false, value === state.selectionAction)));
    action.value = state.selectionAction;
    const valueWrap = document.createElement('span');
    valueWrap.className = 'action-value';
    valueWrap.append(...buildActionValue(state.selectionAction));
    const apply = document.createElement('button');
    apply.type = 'button';
    apply.className = 'button small primary';
    apply.dataset.selectionApply = 'true';
    apply.textContent = 'Apply';
    container.replaceChildren(count, action, valueWrap, apply, clear);
  }

  function buildActionValue(action) {
    const numberInput = (placeholder, label, step = '1') => {
      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.step = step;
      input.placeholder = placeholder;
      input.className = 'action-input';
      input.setAttribute('aria-label', label);
      return input;
    };
    if (action === 'group') {
      const select = document.createElement('select');
      select.className = 'action-group';
      select.setAttribute('aria-label', 'Sales group');
      select.append(new Option('— No group —', ''));
      state.groups.forEach((group) => select.append(new Option(group.name, group.id)));
      if (state.groups.length < MAX_GROUPS) select.append(new Option('+ New group…', '__new'));
      if (state.groups.length) select.value = state.groups[0].id;
      else if (state.groups.length < MAX_GROUPS) select.value = '__new';
      const name = document.createElement('input');
      name.className = 'action-new-group';
      name.maxLength = 40;
      name.placeholder = 'New group name';
      name.setAttribute('aria-label', 'New group name');
      name.hidden = select.value !== '__new';
      return [select, name];
    }
    if (action === 'sales') return [numberInput('Units / yr per item', 'Expected sales per item per year', '0.1')];
    if (action === 'quantity') return [numberInput('Units in stock', 'Stock quantity per item')];
    if (action === 'multiplier') return [numberInput(`× (global ${formatNumber(state.config.multiplier, 2)})`, 'Price multiplier', '0.01')];
    const hint = document.createElement('span');
    hint.className = 'selection-note';
    hint.textContent = {
      exclude: 'Removes the parts from totals, warehouse and dashboard.',
      include: 'Returns excluded parts to the consolidation.',
      clear: 'Groups and exclusions are kept.',
    }[action] || '';
    return [hint];
  }

  function handleSelectionBarChange(event) {
    const bar = event.currentTarget;
    if (event.target.matches('.action-select')) {
      state.selectionAction = event.target.value;
      const valueWrap = bar.querySelector('.action-value');
      valueWrap.replaceChildren(...buildActionValue(state.selectionAction));
      valueWrap.querySelector('input:not([hidden]), select')?.focus();
    } else if (event.target.matches('.action-group')) {
      const name = bar.querySelector('.action-new-group');
      name.hidden = event.target.value !== '__new';
      if (!name.hidden) name.focus();
    }
  }

  function handleSelectionBarClick(event) {
    const bar = event.currentTarget;
    if (event.target.closest('[data-selection-apply]')) applySelectionAction(bar);
    if (event.target.closest('[data-selection-clear]')) {
      if (bar.dataset.context === 'sheet') {
        state.sheetSelection.rows.clear();
        dom.grid.querySelectorAll('tr.row-selected').forEach((row) => row.classList.remove('row-selected'));
        renderSelectionBar(dom.sheetActions, [], 'sheet');
      } else {
        state.selectedKeys.clear();
        renderConsolidated();
        renderSelectionBar(dom.consolidatedActions, [], 'consolidated');
      }
    }
  }

  function applySelectionAction(bar) {
    const keys = bar.dataset.context === 'sheet' ? sheetSelectionKeys() : [...state.selectedKeys];
    if (!keys.length) return;
    const action = state.selectionAction;
    let value = null;
    if (action === 'group') {
      const select = bar.querySelector('.action-group');
      value = select.value;
      if (value === '__new') {
        const groupNameInput = bar.querySelector('.action-new-group');
        const requested = groupNameInput.value;
        const existing = state.groups.find((group) => group.name.toLowerCase() === String(requested || '').trim().toLowerCase());
        if (!existing && state.groups.length >= MAX_GROUPS) {
          showToast(`Up to ${MAX_GROUPS} sales groups are supported so each keeps a distinct color.`, true);
          return;
        }
        const created = pushAndCaptureNewGroup(requested);
        if (!created) return;
        value = created.id;
      }
    } else if (['sales', 'quantity', 'multiplier'].includes(action)) {
      const input = bar.querySelector('.action-input');
      value = toNumber(input.value);
      if (value === null || value < 0) {
        showToast('Enter a value of zero or more first.', true);
        input.focus();
        return;
      }
    }
    // One undo step for the whole bulk edit, so Ctrl+Z reverses the selection rather than each part.
    pushHistory(`${keys.length} part${keys.length === 1 ? '' : 's'} updated`, () => applyItemAction(keys, action, value));
    renderAllSelectionBars();
    const group = action === 'group' ? groupById(value) : null;
    const messages = {
      group: group ? `${keys.length} part${keys.length === 1 ? '' : 's'} assigned to ${group.name}.` : `Sales group removed from ${keys.length} part${keys.length === 1 ? '' : 's'}.`,
      sales: `Expected sales set to ${formatNumber(value, 2)} per year for ${keys.length} part${keys.length === 1 ? '' : 's'}.`,
      quantity: `Stock quantity set to ${formatNumber(value, 2)} for ${keys.length} part${keys.length === 1 ? '' : 's'}.`,
      multiplier: `Price multiplier set to ${formatNumber(value, 2)}× for ${keys.length} part${keys.length === 1 ? '' : 's'}.`,
      exclude: `${keys.length} part${keys.length === 1 ? '' : 's'} excluded. Find them under Show → Excluded parts.`,
      include: `${keys.length} part${keys.length === 1 ? '' : 's'} included again.`,
      clear: `Overrides cleared for ${keys.length} part${keys.length === 1 ? '' : 's'}.`,
    };
    showToast(messages[action]);
  }

  function applyItemAction(keys, action, value) {
    keys.forEach((key) => {
      const item = { group: null, salesPerYear: null, quantity: null, multiplier: null, excluded: false, ...(state.items[key] || {}) };
      if (action === 'group') item.group = value || null;
      if (action === 'sales') item.salesPerYear = value;
      if (action === 'quantity') item.quantity = value;
      if (action === 'multiplier') item.multiplier = value;
      if (action === 'exclude') item.excluded = true;
      if (action === 'include') item.excluded = false;
      if (action === 'clear') {
        item.salesPerYear = null;
        item.quantity = null;
        item.multiplier = null;
      }
      const empty = !item.group && item.salesPerYear === null && item.quantity === null && item.multiplier === null && !item.excluded;
      if (empty) delete state.items[key];
      else state.items[key] = item;
    });
    saveSettingsSoon();
    rebuildConsolidation();
    updateGroupCounts();
  }

  /* ---------- Sales groups ---------- */

  // Creates a group without pushing its own history entry, so "assign to a new group" undoes as one action.
  function pushAndCaptureNewGroup(rawName) {
    return addGroup(rawName, { silent: true });
  }

  function addGroup(rawName, options = {}) {
    const name = String(rawName || '').trim().slice(0, 40);
    if (!name) {
      if (!options.silent) showToast('Enter a name for the new group.', true);
      return null;
    }
    const existing = state.groups.find((group) => group.name.toLowerCase() === name.toLowerCase());
    if (existing) return existing;
    if (state.groups.length >= MAX_GROUPS) {
      if (!options.silent) showToast(`Up to ${MAX_GROUPS} sales groups are supported so each keeps a distinct color.`, true);
      return null;
    }
    const used = new Set(state.groups.map((group) => group.color));
    const group = {
      id: `g-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
      name,
      color: GROUP_COLORS.find((color) => !used.has(color)),
      salesMode: 'turns',
      salesValue: null,
      multiplier: null,
    };
    state.groups.push(group);
    saveSettingsSoon();
    renderGroups();
    renderAllSelectionBars();
    refreshViews();
    if (!options.silent) pushHistory(`Group ${name} added`, () => {});
    return group;
  }

  function removeGroupNow(id) {
    state.groups = state.groups.filter((item) => item.id !== id);
    Object.entries(state.items).forEach(([key, item]) => {
      if (item.group !== id) return;
      item.group = null;
      if (item.salesPerYear === null && item.quantity === null && item.multiplier === null && !item.excluded) delete state.items[key];
    });
    if (state.consolidatedFilter === `group:${id}`) state.consolidatedFilter = 'all';
    saveSettingsSoon();
    renderGroups();
    renderAllSelectionBars();
    rebuildConsolidation();
  }

  function deleteGroup(id) {
    const group = groupById(id);
    if (!group) return;
    pushHistory(`Group ${group.name} deleted`, () => removeGroupNow(id));
    showToast(`Group ${group.name} deleted. Its parts are now ungrouped. Ctrl+Z restores it.`);
  }

  function handleGroupInput(input) {
    const group = groupById(input.dataset.groupId);
    if (!group) return;
    const field = input.dataset.groupField;
    if (field === 'name') {
      const name = input.value.trim().slice(0, 40);
      if (!name) return;
      group.name = name;
      saveSettingsSoon();
      refreshViews();
      return;
    }
    if (field === 'salesMode') group.salesMode = input.value === 'units' ? 'units' : 'turns';
    if (field === 'salesValue' || field === 'multiplier') {
      if (input.value.trim() === '') group[field] = null;
      else {
        const value = readNumberInput(input);
        if (value === undefined) return;
        group[field] = value;
      }
    }
    saveSettingsSoon();
    rebuildConsolidation();
  }

  function renderGroups() {
    if (!state.groups.length) {
      const empty = document.createElement('p');
      empty.className = 'muted';
      empty.textContent = 'No groups yet. Every item uses the default expected sales.';
      dom.groupList.replaceChildren(empty);
      return;
    }
    const rows = state.groups.map((group) => {
      const row = document.createElement('div');
      row.className = 'group-editor';
      row.style.setProperty('--group-color', group.color);
      const head = document.createElement('div');
      head.className = 'group-editor-head';
      const swatch = document.createElement('span');
      swatch.className = 'group-swatch';
      swatch.setAttribute('aria-hidden', 'true');
      const name = document.createElement('input');
      name.value = group.name;
      name.maxLength = 40;
      name.dataset.groupId = group.id;
      name.dataset.groupField = 'name';
      name.setAttribute('aria-label', 'Group name');
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'icon-button';
      remove.dataset.removeGroup = group.id;
      remove.setAttribute('aria-label', `Delete group ${group.name}`);
      remove.textContent = '×';
      head.append(swatch, name, remove);

      const fields = document.createElement('div');
      fields.className = 'group-editor-fields';
      const salesLabel = document.createElement('label');
      salesLabel.textContent = 'Expected sales';
      const salesWrap = document.createElement('span');
      salesWrap.className = 'amount-currency wide-select';
      const salesValue = document.createElement('input');
      salesValue.type = 'number';
      salesValue.min = '0';
      salesValue.step = '0.1';
      salesValue.placeholder = 'default';
      salesValue.dataset.groupId = group.id;
      salesValue.dataset.groupField = 'salesValue';
      salesValue.setAttribute('aria-label', `${group.name} expected sales`);
      writeInputValue(salesValue, group.salesValue);
      const salesMode = document.createElement('select');
      salesMode.dataset.groupId = group.id;
      salesMode.dataset.groupField = 'salesMode';
      salesMode.setAttribute('aria-label', `${group.name} expected sales unit`);
      salesMode.append(new Option('× stock qty / yr', 'turns', false, group.salesMode === 'turns'), new Option('units / item / yr', 'units', false, group.salesMode === 'units'));
      salesWrap.append(salesValue, salesMode);
      salesLabel.append(salesWrap);
      const multiplierLabel = document.createElement('label');
      multiplierLabel.textContent = 'Price multiplier';
      const multiplier = document.createElement('input');
      multiplier.type = 'number';
      multiplier.min = '0';
      multiplier.step = '0.01';
      multiplier.placeholder = `global ${formatNumber(state.config.multiplier, 2)}×`;
      multiplier.dataset.groupId = group.id;
      multiplier.dataset.groupField = 'multiplier';
      multiplier.setAttribute('aria-label', `${group.name} price multiplier`);
      writeInputValue(multiplier, group.multiplier);
      multiplierLabel.append(multiplier);
      const count = document.createElement('small');
      count.className = 'group-count';
      count.dataset.groupCount = group.id;
      fields.append(salesLabel, multiplierLabel, count);
      row.append(head, fields);
      return row;
    });
    dom.groupList.replaceChildren(...rows);
    updateGroupCounts();
  }

  function updateGroupCounts() {
    dom.groupList.querySelectorAll('[data-group-count]').forEach((node) => {
      const id = node.dataset.groupCount;
      const items = state.consolidated.filter((item) => item.groupId === id);
      const units = items.reduce((sum, item) => sum + item.salesPerYear, 0);
      node.textContent = `${items.length} item${items.length === 1 ? '' : 's'} · ${formatNumber(units, 1)} units / yr`;
    });
  }

  /* ---------- Profitability model ---------- */

  function calculateProfitability() {
    const warehouse = calculateWarehouseModel();
    const nok = warehouse.nokToOutput;
    const vatRate = salesVatRate();
    const items = state.consolidated;
    const inventoryUnits = warehouse.inventoryUnits;
    const salesUnits = items.reduce((sum, item) => sum + (item.salesPerYear || 0), 0);
    const unitsPerMonth = salesUnits / 12;
    const convert = (value) => (nok === null ? null : value * nok);
    const storageMonthly = convert(warehouse.storageNok);
    const outtakeMonthly = convert(warehouse.orderHandlingNok + warehouse.ediNok + warehouse.packagingNok);
    const outboundMonthly = convert(warehouse.freightNok);
    const fixedMonthly = convert(warehouse.receivingNok + warehouse.wmsNok);
    const derived = {
      storagePerUnitMonth: storageMonthly === null ? null : inventoryUnits > 0 ? storageMonthly / inventoryUnits : 0,
      outtakePerUnit: outtakeMonthly === null ? null : unitsPerMonth > 0 ? outtakeMonthly / unitsPerMonth : 0,
      outboundPerUnit: outboundMonthly === null ? null : unitsPerMonth > 0 ? outboundMonthly / unitsPerMonth : 0,
    };
    const perUnit = {};
    const manual = {};
    FLOW_OVERRIDE_KEYS.forEach((key) => {
      manual[key] = Number.isFinite(state.flowOverrides[key]);
      perUnit[key] = manual[key] ? state.flowOverrides[key] : derived[key];
    });
    const outboundPerUnit = state.sales.customerPaysOutbound ? 0 : perUnit.outboundPerUnit ?? 0;
    const completeness = Core.assessProfitabilityCompleteness({ warehouseRateMissing: nok === null, items });

    const rows = items.map((item) => {
      const sales = item.salesPerYear || 0;
      const unitPrice = item.unitSalesPrice;
      const unitCost = item.unitLandedCost;
      const revenue = unitPrice === null ? 0 : sales * unitPrice;
      const cogs = unitCost === null ? 0 : sales * unitCost;
      const storage = (perUnit.storagePerUnitMonth ?? 0) * Math.max(0, item.maxQuantity) * 12;
      const handling = sales * ((perUnit.outtakePerUnit ?? 0) + outboundPerUnit);
      const gross = revenue - cogs;
      return {
        item,
        sales,
        unitPrice,
        unitCost,
        complete: completeness.complete && unitPrice !== null && unitCost !== null,
        unitMargin: unitPrice === null || unitCost === null ? null : unitPrice - unitCost,
        revenue,
        cogs,
        gross,
        storage,
        handling,
        net: gross - storage - handling,
      };
    });
    const total = (field) => rows.reduce((sum, row) => sum + row[field], 0);
    const revenue = total('revenue');
    const cogs = total('cogs');
    const gross = revenue - cogs;
    const storage = total('storage');
    const outtake = salesUnits * (perUnit.outtakePerUnit ?? 0);
    const outbound = salesUnits * outboundPerUnit;
    const fixed = (fixedMonthly ?? 0) * 12;
    const net = gross - storage - outtake - outbound - fixed;
    const investment = items.reduce((sum, item) => sum + (item.landedCost ?? 0), 0);
    const monthlyNet = net / 12;
    const paybackMonths = investment > 0 && monthlyNet > 0 ? investment / monthlyNet : null;

    const groupMap = new Map();
    rows.forEach((row) => {
      const id = row.item.groupId || '';
      if (!groupMap.has(id)) {
        const group = groupById(id);
        groupMap.set(id, { id, name: group?.name || 'Ungrouped', color: group?.color || '#898781', items: 0, units: 0, sales: 0, revenue: 0, gross: 0, net: 0 });
      }
      const summary = groupMap.get(id);
      summary.items += 1;
      summary.units += Math.max(0, row.item.maxQuantity);
      summary.sales += row.sales;
      summary.revenue += row.revenue;
      summary.gross += row.gross;
      summary.net += row.net;
    });
    const groupOrder = new Map(state.groups.map((group, index) => [group.id, index]));
    const groups = [...groupMap.values()].sort((a, b) => (groupOrder.get(a.id) ?? 99) - (groupOrder.get(b.id) ?? 99));

    return {
      warehouse,
      rows: rows.sort((a, b) => b.net - a.net),
      groups,
      salesUnits,
      inventoryUnits,
      derived,
      perUnit,
      manual,
      revenue,
      cogs,
      gross,
      storage,
      outtake,
      outbound,
      fixed,
      net,
      investment,
      monthlyNet,
      paybackMonths,
      roi: investment > 0 ? net / investment : null,
      grossMargin: revenue > 0 ? gross / revenue : null,
      outputVat: revenue * vatRate,
      importVat: items.reduce((sum, item) => sum + (item.importVat ?? 0), 0),
      warehouseRateMissing: nok === null,
      complete: completeness.complete,
      missingItemCount: completeness.missingItemCount,
      missingReasons: completeness.reasons,
      averageUnitPrice: salesUnits > 0 ? revenue / salesUnits : null,
    };
  }

  /* ---------- Dashboard ---------- */

  function formatWhole(value) {
    return Number.isFinite(value) ? `${formatNumber(Math.round(value), 0)} ${state.config.outputCurrency}` : '—';
  }

  function formatCompact(value) {
    if (!Number.isFinite(value)) return '—';
    return numberFormat({ notation: 'compact', maximumFractionDigits: Math.abs(value) >= 1e6 ? 2 : 1 }).format(value);
  }

  function kpiTile(label, value, detail, options = {}) {
    const tile = document.createElement('article');
    tile.className = `kpi-tile${options.hero ? ' hero' : ''}`;
    const title = document.createElement('span');
    title.className = 'kpi-label';
    title.textContent = label;
    const number = document.createElement('strong');
    number.className = 'kpi-value';
    number.textContent = value;
    tile.append(title, number);
    if (options.status) {
      const status = document.createElement('span');
      status.className = `kpi-status ${options.status}`;
      status.textContent = options.status === 'good' ? '▲ Profitable' : options.status === 'incomplete' ? 'Estimate incomplete' : '▼ Loss-making';
      tile.append(status);
    }
    if (detail) {
      const small = document.createElement('span');
      small.className = 'kpi-detail';
      small.textContent = detail;
      tile.append(small);
    }
    return tile;
  }

  function setStats(list, pairs) {
    list.replaceChildren(...pairs.flatMap(([term, value]) => {
      const dt = document.createElement('dt');
      dt.textContent = term;
      const dd = document.createElement('dd');
      dd.textContent = value;
      return [dt, dd];
    }));
  }

  function renderDashboard() {
    const model = calculateProfitability();
    state.lastProfitability = model;
    const output = state.config.outputCurrency;
    dom.dashboardSubtitle.textContent = `${model.rows.length} parts from ${enabledSheetCount()} kits · ${formatNumber(model.salesUnits, 1)} expected units sold per year · ${formatNumber(model.inventoryUnits, 0)} units in stock.${model.complete ? '' : ' Charts below show only the costs that can currently be calculated.'}`;
    dom.dashboardBadge.textContent = `${output} · excl. VAT`;
    dom.dashboardIncomplete.classList.toggle('hidden', model.complete);
    dom.dashboardIncomplete.textContent = model.complete ? '' : `Profitability is not final: missing ${model.missingReasons.join(' and ')}. Complete these inputs before relying on profit, ROI or payback.`;
    const logistics = model.storage + model.outtake + model.outbound + model.fixed;
    dom.dashboardKpis.replaceChildren(
      kpiTile('Annual net profit', model.complete ? formatWhole(model.net) : 'Incomplete', model.complete ? `${formatWhole(model.monthlyNet)} per month${model.revenue > 0 ? ` · ${formatPercent(model.net / model.revenue)} of revenue` : ''}` : `${formatWhole(model.net)} known-cost result; do not use as final`, { hero: true, status: model.complete ? (model.net >= 0 ? 'good' : 'critical') : 'incomplete' }),
      kpiTile('Revenue / yr', formatWhole(model.revenue), (salesVatRate() > 0 ? `${formatWhole(model.revenue * (1 + salesVatRate()))} incl. VAT` : 'No VAT included')),
      kpiTile('Gross profit / yr', formatWhole(model.gross), model.grossMargin === null ? 'No expected revenue' : `${formatPercent(model.grossMargin)} gross margin`),
      kpiTile('Warehouse & logistics / yr', formatWhole(logistics), `Storage ${formatWhole(model.storage)} · handling ${formatWhole(model.outtake + model.outbound + model.fixed)}`),
      kpiTile('Payback on stock', model.complete ? (model.paybackMonths === null ? 'Not reached' : `${formatNumber(model.paybackMonths, 1)} months`) : 'Incomplete', `${formatWhole(model.investment)} landed stock${model.complete && model.roi !== null ? ` · ${formatPercent(model.roi)} annual return` : ''}`),
    );
    setStats(dom.vatSummary, [
      ['Output VAT / yr', formatWhole(model.outputVat)],
      ['Import VAT on stock', formatWhole(model.importVat)],
      ['Customer price incl. VAT / yr', formatWhole(model.revenue + model.outputVat)],
    ]);
    updateGroupCounts();
    renderDashboardCharts(model);
    renderGroupTable(model);
    renderDashboardItems(model);
    renderScenarioComparison(model);
  }

  function calculateScenario(model) {
    return Core.projectScenario(model, state.consolidated, state.scenario);
  }

  function renderScenarioComparison(model) {
    const scenario = calculateScenario(model);
    const table = document.createElement('table');
    table.className = 'scenario-table';
    const head = table.createTHead().insertRow();
    ['', 'Current', 'Scenario'].forEach((label) => { const cell = document.createElement('th'); cell.textContent = label; head.append(cell); });
    const body = table.createTBody();
    [
      ['Revenue / yr', formatWhole(model.revenue), formatWhole(scenario.revenue)],
      ['Net profit / yr', model.complete ? formatWhole(model.net) : 'Incomplete', model.complete ? formatWhole(scenario.net) : 'Incomplete'],
      ['Net margin', model.complete && model.revenue > 0 ? formatPercent(model.net / model.revenue) : '—', model.complete && scenario.margin !== null ? formatPercent(scenario.margin) : '—'],
      ['Stock investment', formatWhole(model.investment), formatWhole(scenario.investment)],
      ['Payback', model.complete && model.paybackMonths !== null ? `${formatNumber(model.paybackMonths, 1)} mo` : '—', model.complete && scenario.paybackMonths !== null ? `${formatNumber(scenario.paybackMonths, 1)} mo` : '—'],
    ].forEach(([label, current, changed]) => {
      const row = body.insertRow();
      row.insertCell().textContent = label;
      row.insertCell().textContent = current;
      row.insertCell().textContent = changed;
    });
    dom.scenarioSummary.replaceChildren(table);
  }

  function renderDashboardCharts(model = state.lastProfitability) {
    if (!model) return;
    renderWaterfallChart(dom.chartWaterfall, model);
    renderCumulativeChart(dom.chartCumulative, model);
    renderGroupChart(dom.chartGroups, model);
  }

  function chartWidth(container, minimum = 300) {
    return Math.max(minimum, Math.floor(container.clientWidth || container.parentElement?.clientWidth || 560));
  }

  function niceTicks(min, max, count = 5) {
    if (min === max) {
      const pad = Math.abs(min) || 1;
      min -= pad;
      max += pad;
    }
    const rawStep = (max - min) / count;
    const magnitude = 10 ** Math.floor(Math.log10(rawStep));
    const step = [1, 2, 2.5, 5, 10].map((factor) => factor * magnitude).find((candidate) => candidate >= rawStep) || rawStep;
    const start = Math.floor(min / step) * step;
    const end = Math.ceil(max / step) * step;
    const ticks = [];
    for (let value = start; value <= end + step / 2; value += step) ticks.push(Math.abs(value) < step / 1e6 ? 0 : value);
    return { ticks, min: start, max: end };
  }

  // Bar path with a 4px rounded data end and a square end at the baseline.
  function horizontalBarPath(x0, x1, y, height) {
    const left = Math.min(x0, x1);
    const right = Math.max(x0, x1);
    const radius = Math.min(4, (right - left) / 2, height / 2);
    if (right - left < 0.5) return `M${left},${y} v${height} h0.5 v${-height} Z`;
    return x1 >= x0
      ? `M${left},${y} H${right - radius} Q${right},${y} ${right},${y + radius} V${y + height - radius} Q${right},${y + height} ${right - radius},${y + height} H${left} Z`
      : `M${right},${y} H${left + radius} Q${left},${y} ${left},${y + radius} V${y + height - radius} Q${left},${y + height} ${left + radius},${y + height} H${right} Z`;
  }

  function tipAttributes(value, label) {
    return `data-tip-value="${escapeMarkup(value)}" data-tip-label="${escapeMarkup(label)}" tabindex="0"`;
  }

  // Horizontal profit bridge: step names left, floating bars in the middle, values in their own column.
  function renderWaterfallChart(container, model) {
    const output = state.config.outputCurrency;
    const steps = [
      { label: 'Revenue', value: model.revenue, total: true },
      { label: 'Landed cost of goods', value: -model.cogs },
      { label: 'Storage', value: -model.storage },
      { label: 'Outtake & orders', value: -model.outtake },
      ...(state.sales.customerPaysOutbound ? [] : [{ label: 'Outbound freight', value: -model.outbound }]),
      { label: 'Receiving & WMS', value: -model.fixed },
      { label: 'Net profit', value: model.net, total: true },
    ];
    let level = 0;
    steps.forEach((step) => {
      step.from = step.total ? 0 : level;
      step.to = step.total ? step.value : level + step.value;
      level = step.to;
    });
    const width = chartWidth(container);
    const rowHeight = 32;
    const barHeight = 18;
    const longestLabel = Math.max(...steps.map((step) => step.label.length));
    const margin = { top: 6, right: 78, bottom: 22, left: Math.min(width * 0.42, Math.max(90, longestLabel * 6.4 + 14)) };
    const height = margin.top + margin.bottom + steps.length * rowHeight;
    const plotWidth = width - margin.left - margin.right;
    const values = steps.flatMap((step) => [step.from, step.to]);
    const scale = niceTicks(Math.min(0, ...values), Math.max(0, ...values), 4);
    const x = (value) => margin.left + ((value - scale.min) / (scale.max - scale.min)) * plotWidth;
    const parts = [];
    scale.ticks.forEach((tick) => {
      parts.push(`<line class="viz-grid" x1="${x(tick).toFixed(1)}" x2="${x(tick).toFixed(1)}" y1="${margin.top}" y2="${(height - margin.bottom).toFixed(1)}"/>`);
      parts.push(`<text class="viz-tick" x="${x(tick).toFixed(1)}" y="${height - 6}" text-anchor="middle">${escapeMarkup(formatCompact(tick))}</text>`);
    });
    parts.push(`<line class="viz-axis" x1="${x(0).toFixed(1)}" x2="${x(0).toFixed(1)}" y1="${margin.top}" y2="${(height - margin.bottom).toFixed(1)}"/>`);
    steps.forEach((step, index) => {
      const top = margin.top + index * rowHeight;
      const barY = top + (rowHeight - barHeight) / 2;
      const positive = step.value >= 0;
      if (index < steps.length - 1) {
        parts.push(`<line class="viz-connector" x1="${x(step.to).toFixed(1)}" x2="${x(step.to).toFixed(1)}" y1="${(barY + barHeight).toFixed(1)}" y2="${(barY + rowHeight).toFixed(1)}"/>`);
      }
      parts.push(`<path class="${positive ? 'viz-positive' : 'viz-negative'}" d="${horizontalBarPath(x(step.from), x(step.to), barY, barHeight)}"/>`);
      parts.push(`<text class="viz-category${step.total ? ' viz-total-label' : ''}" x="${margin.left - 10}" y="${(barY + barHeight / 2 + 4).toFixed(1)}" text-anchor="end">${escapeMarkup(step.label)}</text>`);
      parts.push(`<text class="viz-value" x="${width - 4}" y="${(barY + barHeight / 2 + 4).toFixed(1)}" text-anchor="end">${escapeMarkup(formatCompact(step.value))}</text>`);
      parts.push(`<rect class="viz-hit" x="0" y="${top}" width="${width}" height="${rowHeight}" ${tipAttributes(formatWhole(step.value), `${step.label} per year`)}/>`);
    });
    const legend = `<div class="viz-legend"><span><i class="viz-swatch viz-positive"></i>Revenue &amp; profit</span><span><i class="viz-swatch viz-negative"></i>Costs &amp; loss</span></div>`;
    container.innerHTML = `${legend}<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="Annual profit bridge from revenue ${escapeMarkup(formatWhole(model.revenue))} to net profit ${escapeMarkup(formatWhole(model.net))}">${parts.join('')}</svg>`;

    const table = document.createElement('table');
    table.className = 'mini-table';
    const head = table.createTHead().insertRow();
    ['Step', `Amount / yr (${output})`].forEach((text) => { const th = document.createElement('th'); th.textContent = text; head.append(th); });
    const body = table.createTBody();
    steps.forEach((step) => {
      const row = body.insertRow();
      row.insertCell().textContent = step.label;
      const cell = row.insertCell();
      cell.textContent = formatWhole(step.value);
      cell.className = 'number';
    });
    dom.waterfallTable.replaceChildren(table);
  }

  function renderCumulativeChart(container, model) {
    const horizon = Math.round(clamp(state.sales.horizonMonths || 24, 6, 120));
    const series = Array.from({ length: horizon + 1 }, (_, month) => ({ month, value: -model.investment + month * model.monthlyNet }));
    const width = chartWidth(container);
    const height = 270;
    const margin = { top: 22, right: 64, bottom: 34, left: 58 };
    const plotWidth = width - margin.left - margin.right;
    const plotHeight = height - margin.top - margin.bottom;
    const values = series.map((point) => point.value);
    const scale = niceTicks(Math.min(0, ...values), Math.max(0, ...values));
    const x = (month) => margin.left + (month / horizon) * plotWidth;
    const y = (value) => margin.top + plotHeight - ((value - scale.min) / (scale.max - scale.min)) * plotHeight;
    const parts = [];
    scale.ticks.forEach((tick) => {
      parts.push(`<line class="viz-grid" x1="${margin.left}" x2="${width - margin.right}" y1="${y(tick).toFixed(1)}" y2="${y(tick).toFixed(1)}"/>`);
      parts.push(`<text class="viz-tick" x="${margin.left - 8}" y="${(y(tick) + 4).toFixed(1)}" text-anchor="end">${escapeMarkup(formatCompact(tick))}</text>`);
    });
    const tickStep = horizon <= 12 ? 3 : horizon <= 36 ? 6 : 12;
    for (let month = 0; month <= horizon; month += tickStep) {
      parts.push(`<text class="viz-tick" x="${x(month).toFixed(1)}" y="${height - margin.bottom + 18}" text-anchor="middle">${month}</text>`);
    }
    parts.push(`<text class="viz-tick" x="${(width - margin.right).toFixed(1)}" y="${height - 4}" text-anchor="end">months</text>`);
    parts.push(`<line class="viz-axis" x1="${margin.left}" x2="${width - margin.right}" y1="${y(0).toFixed(1)}" y2="${y(0).toFixed(1)}"/>`);
    const line = series.map((point, index) => `${index ? 'L' : 'M'}${x(point.month).toFixed(1)},${y(point.value).toFixed(1)}`).join(' ');
    parts.push(`<path class="viz-area" d="${line} L${x(horizon).toFixed(1)},${y(0).toFixed(1)} L${x(0).toFixed(1)},${y(0).toFixed(1)} Z"/>`);
    parts.push(`<path class="viz-line" d="${line}"/>`);
    if (model.paybackMonths !== null && model.paybackMonths <= horizon) {
      const bx = x(model.paybackMonths);
      parts.push(`<line class="viz-marker-line" x1="${bx.toFixed(1)}" x2="${bx.toFixed(1)}" y1="${margin.top}" y2="${(height - margin.bottom).toFixed(1)}"/>`);
      parts.push(`<circle class="viz-dot" cx="${bx.toFixed(1)}" cy="${y(0).toFixed(1)}" r="4.5"/>`);
      const anchor = bx > width - margin.right - 110 ? 'end' : 'start';
      parts.push(`<text class="viz-value" x="${(bx + (anchor === 'start' ? 7 : -7)).toFixed(1)}" y="${(margin.top + 10).toFixed(1)}" text-anchor="${anchor}">Break-even · month ${escapeMarkup(formatNumber(model.paybackMonths, 1))}</text>`);
    }
    const last = series[series.length - 1];
    parts.push(`<circle class="viz-dot" cx="${x(last.month).toFixed(1)}" cy="${y(last.value).toFixed(1)}" r="4.5"/>`);
    parts.push(`<text class="viz-value" x="${(x(last.month) + 8).toFixed(1)}" y="${(y(last.value) + 4).toFixed(1)}">${escapeMarkup(formatCompact(last.value))}</text>`);
    parts.push(`<line class="viz-crosshair" x1="0" x2="0" y1="${margin.top}" y2="${(height - margin.bottom).toFixed(1)}" visibility="hidden"/>`);
    parts.push(`<circle class="viz-dot viz-hover-dot" cx="0" cy="0" r="4.5" visibility="hidden"/>`);
    parts.push(`<rect class="viz-hit viz-plot-hit" x="${margin.left}" y="${margin.top}" width="${plotWidth}" height="${plotHeight}" tabindex="0" aria-label="Cumulative net position by month"/>`);
    container.innerHTML = `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="Cumulative net position over ${horizon} months, starting at ${escapeMarkup(formatWhole(-model.investment))}">${parts.join('')}</svg>`;

    const svg = container.querySelector('svg');
    const crosshair = svg.querySelector('.viz-crosshair');
    const dot = svg.querySelector('.viz-hover-dot');
    const hit = svg.querySelector('.viz-plot-hit');
    const showMonth = (month, clientX, clientY) => {
      const point = series[clamp(Math.round(month), 0, horizon)];
      crosshair.setAttribute('x1', x(point.month));
      crosshair.setAttribute('x2', x(point.month));
      crosshair.setAttribute('visibility', 'visible');
      dot.setAttribute('cx', x(point.month));
      dot.setAttribute('cy', y(point.value));
      dot.setAttribute('visibility', 'visible');
      showChartTooltip(formatWhole(point.value), `Month ${point.month} · net position`, clientX, clientY);
    };
    const hide = () => {
      crosshair.setAttribute('visibility', 'hidden');
      dot.setAttribute('visibility', 'hidden');
      hideChartTooltip();
    };
    hit.addEventListener('pointermove', (event) => {
      const rect = svg.getBoundingClientRect();
      const localX = (event.clientX - rect.left) * (width / rect.width);
      showMonth(((localX - margin.left) / plotWidth) * horizon, event.clientX, event.clientY);
    });
    hit.addEventListener('pointerleave', hide);
    hit.addEventListener('blur', hide);
    let focusMonth = 0;
    hit.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowRight' && event.key !== 'ArrowLeft') return;
      event.preventDefault();
      focusMonth = clamp(focusMonth + (event.key === 'ArrowRight' ? 1 : -1), 0, horizon);
      const rect = svg.getBoundingClientRect();
      showMonth(focusMonth, rect.left + (x(focusMonth) / width) * rect.width, rect.top + (y(series[focusMonth].value) / height) * rect.height);
    });

    const table = document.createElement('table');
    table.className = 'mini-table';
    const head = table.createTHead().insertRow();
    ['Month', `Net position (${state.config.outputCurrency})`].forEach((text) => { const th = document.createElement('th'); th.textContent = text; head.append(th); });
    const body = table.createTBody();
    series.filter((point) => point.month % 3 === 0 || point.month === horizon).forEach((point) => {
      const row = body.insertRow();
      row.insertCell().textContent = String(point.month);
      const cell = row.insertCell();
      cell.textContent = formatWhole(point.value);
      cell.className = 'number';
    });
    dom.cumulativeTable.replaceChildren(table);
  }

  function renderGroupChart(container, model) {
    const groups = model.groups;
    if (groups.length < 2) {
      container.innerHTML = '';
      const hint = document.createElement('p');
      hint.className = 'chart-empty';
      hint.textContent = state.groups.length
        ? 'Assign parts to more than one group to compare them here.'
        : 'All parts are ungrouped. Select parts in Consolidated or a kit tab and choose “Assign to sales group” to compare groups here.';
      container.append(hint);
      return;
    }
    const width = chartWidth(container);
    const rowHeight = 34;
    const barHeight = 18;
    const margin = { top: 8, right: 76, bottom: 22, left: Math.min(170, Math.max(110, width * 0.24)) };
    const height = margin.top + margin.bottom + groups.length * rowHeight;
    const plotWidth = width - margin.left - margin.right;
    const values = groups.map((group) => group.net);
    const scale = niceTicks(Math.min(0, ...values), Math.max(0, ...values), 4);
    const x = (value) => margin.left + ((value - scale.min) / (scale.max - scale.min)) * plotWidth;
    const parts = [];
    scale.ticks.forEach((tick) => {
      parts.push(`<line class="viz-grid" x1="${x(tick).toFixed(1)}" x2="${x(tick).toFixed(1)}" y1="${margin.top}" y2="${(height - margin.bottom).toFixed(1)}"/>`);
      parts.push(`<text class="viz-tick" x="${x(tick).toFixed(1)}" y="${height - 6}" text-anchor="middle">${escapeMarkup(formatCompact(tick))}</text>`);
    });
    parts.push(`<line class="viz-axis" x1="${x(0).toFixed(1)}" x2="${x(0).toFixed(1)}" y1="${margin.top}" y2="${(height - margin.bottom).toFixed(1)}"/>`);
    groups.forEach((group, index) => {
      const top = margin.top + index * rowHeight;
      const barY = top + (rowHeight - barHeight) / 2;
      const name = group.name.length > 22 ? `${group.name.slice(0, 21)}…` : group.name;
      parts.push(`<text class="viz-category" x="${margin.left - 10}" y="${(barY + barHeight / 2 + 4).toFixed(1)}" text-anchor="end">${escapeMarkup(name)}</text>`);
      parts.push(`<path d="${horizontalBarPath(x(0), x(group.net), barY, barHeight)}" fill="${escapeMarkup(group.color)}"/>`);
      const end = x(group.net);
      const anchor = group.net >= 0 ? 'start' : 'end';
      parts.push(`<text class="viz-value" x="${(end + (group.net >= 0 ? 6 : -6)).toFixed(1)}" y="${(barY + barHeight / 2 + 4).toFixed(1)}" text-anchor="${anchor}">${escapeMarkup(formatCompact(group.net))}</text>`);
      parts.push(`<rect class="viz-hit" x="0" y="${top}" width="${width}" height="${rowHeight}" ${tipAttributes(formatWhole(group.net), `${group.name} · net profit / yr · ${group.items} items · revenue ${formatWhole(group.revenue)}`)}/>`);
    });
    container.innerHTML = `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="Annual net profit by sales group">${parts.join('')}</svg>`;
  }

  function renderGroupTable(model) {
    const output = state.config.outputCurrency;
    const table = document.createElement('table');
    table.className = 'mini-table';
    const head = table.createTHead().insertRow();
    ['Sales group', 'Items', 'Stock units', 'Sales units / yr', `Revenue / yr (${output})`, `Gross profit / yr`, `Net profit / yr`, 'Gross margin'].forEach((text, index) => {
      const th = document.createElement('th');
      th.textContent = text;
      if (index) th.className = 'number';
      head.append(th);
    });
    const body = table.createTBody();
    model.groups.forEach((group) => {
      const row = body.insertRow();
      const nameCell = row.insertCell();
      const chip = document.createElement('span');
      chip.className = 'group-chip';
      chip.style.setProperty('--group-color', group.color);
      chip.textContent = group.name;
      chip.translate = false;
      nameCell.append(chip);
      [formatNumber(group.items, 0), formatNumber(group.units, 0), formatNumber(group.sales, 1), formatWhole(group.revenue), formatWhole(group.gross), formatWhole(group.net), group.revenue > 0 ? formatPercent(group.gross / group.revenue) : '—']
        .forEach((value) => { const cell = row.insertCell(); cell.textContent = value; cell.className = 'number'; });
    });
    dom.groupTable.replaceChildren(table);
  }

  function renderDashboardItems(model) {
    const output = state.config.outputCurrency;
    const table = document.createElement('table');
    table.className = 'parts-table dashboard-table';
    const headers = ['Part', 'Description', 'Sales group', 'Stock qty', 'Sales / yr', `Unit price (${output})`, 'Unit landed cost', 'Unit margin', 'Margin %', 'Revenue / yr', 'Gross profit / yr', 'Storage / yr', 'Outtake & freight / yr', 'Net profit / yr'];
    const head = table.createTHead().insertRow();
    headers.forEach((text, index) => { const th = document.createElement('th'); th.textContent = text; if (index > 2) th.className = 'number'; head.append(th); });
    const body = table.createTBody();
    model.rows.forEach((row) => {
      const tr = body.insertRow();
      if (row.item.common) tr.className = 'common';
      const partCell = tr.insertCell();
      partCell.textContent = row.item.part;
      partCell.translate = false;
      const descriptionCell = tr.insertCell();
      descriptionCell.textContent = row.item.description;
      descriptionCell.translate = false;
      const groupCell = tr.insertCell();
      if (row.item.groupId) {
        const chip = document.createElement('span');
        chip.className = 'group-chip';
        chip.style.setProperty('--group-color', row.item.groupColor);
        chip.textContent = row.item.groupName;
        chip.translate = false;
        groupCell.append(chip);
      } else {
        groupCell.textContent = '—';
      }
      [
        formatNumber(row.item.maxQuantity, 2), formatNumber(row.sales, 2), formatNullableNumber(row.unitPrice, 2), formatNullableNumber(row.unitCost, 2),
        formatNullableNumber(row.unitMargin, 2), row.unitPrice ? formatPercent(row.unitMargin / row.unitPrice) : '—',
        formatNumber(row.revenue, 0), formatNumber(row.gross, 0), formatNumber(row.storage, 0), formatNumber(row.handling, 0), formatNumber(row.net, 0),
      ].forEach((value, index, list) => {
        const cell = tr.insertCell();
        cell.textContent = value;
        cell.className = 'number';
        if (index === list.length - 1) cell.classList.add('strong-cell');
      });
    });
    dom.dashboardItems.replaceChildren(table);
  }

  /* ---------- Chart tooltip ---------- */

  function bindChartTooltip() {
    const show = (event) => {
      const target = event.target.closest?.('[data-tip-value]');
      if (!target) return;
      const rect = target.getBoundingClientRect();
      const x = event.clientX ?? rect.left + rect.width / 2;
      const y = event.clientY ?? rect.top;
      showChartTooltip(target.dataset.tipValue, target.dataset.tipLabel, x || rect.left + rect.width / 2, y || rect.top);
    };
    document.addEventListener('pointermove', (event) => {
      if (event.target.closest?.('[data-tip-value]')) show(event);
    });
    document.addEventListener('pointerout', (event) => {
      if (event.target.closest?.('[data-tip-value]') && !event.relatedTarget?.closest?.('[data-tip-value]')) hideChartTooltip();
    });
    document.addEventListener('focusin', (event) => {
      const target = event.target.closest?.('[data-tip-value]');
      if (!target) return;
      const rect = target.getBoundingClientRect();
      showChartTooltip(target.dataset.tipValue, target.dataset.tipLabel, rect.left + rect.width / 2, rect.top + 20);
    });
    document.addEventListener('focusout', (event) => { if (event.target.closest?.('[data-tip-value]')) hideChartTooltip(); });
  }

  function showChartTooltip(value, label, clientX, clientY) {
    const tip = dom.chartTooltip;
    const strong = document.createElement('strong');
    strong.textContent = value;
    const span = document.createElement('span');
    span.textContent = label;
    tip.replaceChildren(strong, span);
    tip.hidden = false;
    const { width, height } = tip.getBoundingClientRect();
    const left = clamp(clientX + 14, 8, window.innerWidth - width - 8);
    const top = clientY - height - 12 < 8 ? clientY + 16 : clientY - height - 12;
    tip.style.transform = `translate(${Math.round(left)}px, ${Math.round(top)}px)`;
  }

  function hideChartTooltip() {
    if (dom.chartTooltip) dom.chartTooltip.hidden = true;
  }

  /* ---------- Factory-to-customer flow ---------- */

  function renderFlow() {
    const model = calculateProfitability();
    const warehouse = model.warehouse;
    const items = state.consolidated;
    const sum = (field) => items.reduce((total, item) => total + (item[field] ?? 0), 0);
    const destination = destinationConfig();
    const customer = customerConfig();
    dom.flowCustomerDestination.textContent = `Customer in ${customer.name}`;
    dom.destinationMapTitle.textContent = isCrossBorderSale() ? `${destination.name} → ${customer.name}` : destination.name;
    dom.flowCustomerNote.textContent = isCrossBorderSale()
      ? `Sales from a ${destination.name} warehouse to customers in ${customer.name} cross a border. Export documents and the customer-side VAT are not calculated yet.`
      : '';
    dom.flowCustomerNote.classList.toggle('hidden', !isCrossBorderSale());
    renderFlowLanes();

    const palletKey = `${state.warehouse.palletType}${state.warehouse.palletHeight}`;
    dom.flowPalletRate.dataset.bind = `warehouse.rates.${palletKey}`;
    dom.flowPalletLabel.textContent = `${state.warehouse.palletType === 'eu' ? 'EU' : 'Sea'} pallet ≤${state.warehouse.palletHeight} / mo`;
    dom.flowWarehouseTotal.textContent = `${formatNok(warehouse.monthlyNok)} · ${formatWhole(warehouse.monthlyTotal)}`;
    setStats(dom.flowWarehouseStats, [
      ['Locations', `${warehouse.shelfLocations} shelf · ${warehouse.drawerLocations} drawer · ${warehouse.pallets} pallet`],
      ['Per month', formatWhole(warehouse.monthlyTotal)],
      ['Per year', formatWhole(warehouse.annualTotal)],
    ]);

    const active = document.activeElement;
    FLOW_OVERRIDE_KEYS.forEach((key) => {
      const input = dom.flowPerItem.querySelector(`[data-override="${key}"]`);
      const status = dom.flowPerItem.querySelector(`[data-override-state="${key}"]`);
      const value = model.perUnit[key];
      if (input !== active) input.value = Number.isFinite(value) ? String(Math.round(value * 100) / 100) : '';
      status.textContent = model.manual[key] ? 'manual' : Number.isFinite(value) ? state.config.outputCurrency : 'rate missing';
      status.classList.toggle('manual', model.manual[key]);
    });
    dom.flowResetOverrides.disabled = !FLOW_OVERRIDE_KEYS.some((key) => model.manual[key]);
    renderFlowLock();
    if (state.flowMode === 'map') renderMaps();

    setStats(dom.flowCustomerStats, [
      ['Avg. price / unit', model.averageUnitPrice === null ? '—' : `${formatWhole(model.averageUnitPrice)} excl. VAT`],
      ...(salesVatRate() > 0
        ? [['incl. VAT', model.averageUnitPrice === null ? '—' : formatWhole(model.averageUnitPrice * (1 + salesVatRate()))]]
        : [['VAT', 'Not included']]),
      ['Sales / yr', `${formatNumber(model.salesUnits, 0)} units · ${formatWhole(model.revenue)}`],
      ['Net profit / yr', formatWhole(model.net)],
    ]);
  }

  function renderFlowLock() {
    const unlocked = state.flowUnlocked;
    dom.flowPerItem.classList.toggle('locked', !unlocked);
    dom.flowLock.setAttribute('aria-pressed', String(unlocked));
    dom.flowLock.querySelector('.lock-label').textContent = unlocked ? 'Unlocked' : 'Locked';
    dom.flowLock.title = unlocked ? 'Lock the per-item values' : 'Unlock to edit the per-item values';
    dom.flowPerItem.querySelectorAll('[data-override]').forEach((input) => { input.readOnly = !unlocked; });
  }

  /* ---------- Customs & VAT guide ---------- */

  function originName(id = state.supplier.country) {
    const english = ORIGIN_PRESETS[id]?.name || state.map.names?.get(id) || state.supplier.countryName || 'Selected country';
    return I18n && state.language === 'sv' ? I18n.countryName('sv', id, english) : english;
  }

  function destinationConfig() {
    return DESTINATIONS[state.warehouse.country] || DESTINATIONS[WAREHOUSE_DEFAULTS.country];
  }

  // The country the goods are sold into. It is separate from the warehouse country: a warehouse in one
  // country can serve customers in another.
  function customerConfig() {
    return DESTINATIONS[state.customer.country] || DESTINATIONS['578'];
  }

  // VAT is off (rate 0) unless the rate is unlocked and set; it only applies to sales countries set to "price".
  function salesVatRate() {
    return Core.salesVatRate(state.config.vatRate, state.vat.byCountry, state.customer.country);
  }

  function vatTreatment(countryId = state.customer.country) {
    return state.vat.byCountry[countryId] === 'price' ? 'price' : 'none';
  }

  function renderVatControls() {
    const unlocked = state.vat.unlocked;
    document.querySelectorAll('[data-vat-lock]').forEach((button) => {
      button.setAttribute('aria-pressed', String(unlocked));
      button.title = unlocked ? 'Lock VAT' : 'Unlock to include VAT';
      const label = button.querySelector('.lock-label');
      if (label) label.textContent = unlocked ? 'Unlocked' : 'Locked';
    });
    document.querySelectorAll('[data-vat-field]').forEach((input) => { input.readOnly = !unlocked; });
    if (dom.vatCountries) {
      const ids = Object.keys(DESTINATIONS);
      const signature = `${state.language}|${ids.join(',')}`;
      if (dom.vatCountries.dataset.signature !== signature) {
        dom.vatCountries.dataset.signature = signature;
        dom.vatCountries.replaceChildren(...ids.map((id) => {
          const label = document.createElement('label');
          label.append(document.createTextNode(originName(id)));
          const select = document.createElement('select');
          select.dataset.vatCountry = id;
          select.setAttribute('aria-label', originName(id));
          VAT_TREATMENT_KEYS.forEach((key) => select.append(new Option(VAT_LABELS[key], key)));
          label.append(select);
          return label;
        }));
      }
      dom.vatCountries.querySelectorAll('select').forEach((select) => {
        select.value = vatTreatment(select.dataset.vatCountry);
        select.disabled = !unlocked;
      });
    }
    if (dom.vatState) {
      dom.vatState.textContent = state.config.vatRate > 0
        ? `VAT rate ${formatPercent(salesVatRate())} applies to sales into countries set to "VAT on the price after margin".`
        : 'VAT is off: the project calculates without VAT (0 %). Unlock the rate to include it.';
    }
    if (dom.vatRoutes) {
      const customer = customerConfig();
      setStats(dom.vatRoutes, guideOrigins().map((origin) => [`${originName(origin)} → ${customer.name}`, VAT_LABELS[vatTreatment()]]));
    }
  }

  function isCrossBorderSale() {
    return state.customer.country !== state.warehouse.country;
  }

  function populateCustomerSelects() {
    const ids = Object.keys(DESTINATIONS);
    document.querySelectorAll('[data-customer-select]').forEach((select) => {
      const signature = `${state.language}|${ids.join(',')}`;
      if (select.dataset.signature !== signature) {
        select.dataset.signature = signature;
        select.replaceChildren(...ids.map((id) => new Option(originName(id), id)));
      }
      if (select !== document.activeElement) select.value = state.customer.country;
    });
  }

  function originAgreement(id) {
    if (id === '578') return 'norway';
    if (EU_MEMBERS.has(id)) return 'eu';
    if (EFTA_MEMBERS.has(id)) return 'efta';
    if (id === '840') return 'usa';
    return 'other';
  }

  const AGREEMENT_TEXT = {
    intraeu: {
      badge: 'Intra-EU movement',
      duty: 'Goods moving from another EU member state to Sweden are not imported at an external EU border, so no customs duty or import declaration is normally made in Sweden. Swedish acquisition VAT is instead handled in the VAT return.',
    },
    eu: {
      badge: 'EU · EEA agreement',
      duty: 'Free trade through the EEA agreement. Dutiable goods (such as clothing or food) are 0% only when the shipment carries valid proof of origin: an EUR.1 certificate or a valid invoice declaration showing the goods were made in the EU.',
    },
    efta: {
      badge: 'EFTA member',
      duty: 'Free trade through the EFTA convention. Dutiable goods are 0% only when the shipment carries valid proof of origin: an EUR.1 certificate or a valid invoice declaration.',
    },
    usa: {
      badge: 'No broad free trade agreement',
      duty: 'Norway has no broad free trade agreement with the USA, so ordinary Norwegian rates apply. Dutiable goods such as textiles pay duty, but machinery, electronics and most consumer goods already have 0% base duty.',
    },
    other: {
      badge: 'Ordinary rates unless an agreement applies',
      duty: 'Ordinary Norwegian rates apply unless a trade agreement covers this origin. Most industrial and technical goods are duty-free in Norway regardless of origin.',
    },
    norway: {
      badge: 'Domestic',
      duty: 'Goods bought inside Norway are not imported, so no customs duty or import VAT applies.',
    },
  };

  const CUSTOMS_STRATEGIES = [
    { key: 'tollager', title: 'Use a customs warehouse (Tollager)', text: 'Store dutiable goods uncleared at a warehouse partner with a Tollager permit. Duty is paid only when a part leaves for a Norwegian customer; parts shipped on out of Norway (for example to Sweden) pay no Norwegian duty.', origins: ['usa', 'other'] },
    { key: 't1', title: 'Truck through the EU under T1 transit', text: 'Goods trucked from Bulgaria or Switzerland through the EU and Sweden must travel under a T1 transit document. If the forwarder clears them into the EU by mistake, you pay EU duty first and then Norwegian duty and fees at the border.', origins: ['eu', 'efta'] },
    { key: 'tollkreditt', title: 'Apply for Tollkreditt', text: 'If duty or special fees do arise, apply for Tollkreditt with Skatteetaten. Charges are collected on one monthly invoice instead of the forwarder paying at the border and billing disbursement fees.', origins: ['eu', 'efta', 'usa', 'other'] },
    { key: 'incoterms', title: 'Buy DAP or FCA, not DDP', text: 'Buy on DAP (Delivered at Place) or FCA (Free Carrier). Avoid DDP (Delivered Duty Paid): the supplier then handles Norwegian customs, you lose control of the declaration, and the import VAT deduction gets lost or complicated.', origins: ['eu', 'efta', 'usa', 'other'] },
  ];

  function buildCustomsGuide(countryId = state.supplier.country, { compact = false } = {}) {
    const destination = destinationConfig();
    let agreement = originAgreement(countryId);
    let info = AGREEMENT_TEXT[agreement];
    if (state.warehouse.country === '752') {
      if (countryId === '752') info = { badge: 'Domestic', duty: 'Goods bought and warehoused in Sweden do not cross a customs border, so no customs duty or import VAT applies.' };
      else if (EU_MEMBERS.has(countryId)) { agreement = 'intraeu'; info = AGREEMENT_TEXT.intraeu; }
      else if (countryId === '578' || EFTA_MEMBERS.has(countryId)) {
        agreement = 'efta';
        info = { badge: 'Non-EU import · preferential origin may apply', duty: 'Imports from Norway or Switzerland require EU customs clearance. Preferential duty can apply when the goods meet the relevant origin rules and valid proof of origin accompanies the shipment; otherwise the EU Common Customs Tariff applies.' };
      } else if (countryId === '840') {
        info = { badge: 'Non-EU import', duty: 'Imports from the United States require EU customs clearance. Duty depends on the commodity code and origin under the EU Common Customs Tariff.' };
      } else {
        info = { badge: 'Non-EU import', duty: 'Imports into Sweden from outside the EU require customs clearance. Duty depends on the commodity code, customs value, origin and any applicable trade agreement.' };
      }
    }
    const fragment = document.createDocumentFragment();
    const head = document.createElement('div');
    head.className = 'guide-head';
    const title = document.createElement('strong');
    title.textContent = `${originName(countryId)} → ${destination.name}`;
    const badge = document.createElement('span');
    badge.className = `guide-badge ${agreement}`;
    badge.textContent = info.badge;
    head.append(title, badge);
    fragment.append(head);

    const section = (heading, text) => {
      const block = document.createElement('div');
      block.className = 'guide-block';
      const h = document.createElement('h5');
      h.textContent = heading;
      const p = document.createElement('p');
      p.textContent = text;
      block.append(h, p);
      return block;
    };
    const norwayDetail = ' In Norway most industrial and technical goods are duty-free wherever they are made; confirm the HS code and proof-of-origin requirements with the forwarder.';
    fragment.append(section('Customs duty', `${info.duty}${state.warehouse.country === '578' ? norwayDetail : ' Confirm the CN/HS commodity code, customs value and origin documentation with the customs representative.'}`));
    if (countryId !== state.warehouse.country && agreement !== 'intraeu') {
      const fromCountry = state.consolidated.filter((item) => item.originCountry === countryId);
      const importVat = (fromCountry.length ? fromCountry : state.consolidated).reduce((sum, item) => sum + (item.importVat ?? 0), 0);
      const vatText = state.warehouse.country === '752'
        ? `Swedish import VAT is normally reported to Skatteverket by a VAT-registered importer and is calculated from the customs value plus duty and certain ancillary costs — about ${formatWhole(importVat)} for this scenario.`
        : `Normally 25% (15% on food), calculated on the purchase value plus freight and any duty — about ${formatWhole(importVat)} for this order. A Norwegian VAT-registered importer normally reports import VAT in the MVA return.`;
      fragment.append(section('Import VAT', vatText));
      const strategies = state.warehouse.country === '578' ? CUSTOMS_STRATEGIES.filter((strategy) => strategy.origins.includes(agreement)) : [];
      const list = document.createElement(compact ? 'ul' : 'div');
      list.className = compact ? 'guide-list' : 'guide-strategies';
      strategies.forEach((strategy) => {
        if (compact) {
          const item = document.createElement('li');
          const strong = document.createElement('strong');
          strong.textContent = `${strategy.title}. `;
          item.append(strong, document.createTextNode(strategy.text));
          list.append(item);
        } else {
          const details = document.createElement('details');
          const summary = document.createElement('summary');
          summary.textContent = strategy.title;
          const p = document.createElement('p');
          p.textContent = strategy.text;
          details.append(summary, p);
          list.append(details);
        }
      });
      if (strategies.length) {
        const h = document.createElement('h5');
        h.textContent = 'Avoid unnecessary cost and border stops';
        fragment.append(h, list);
      }
    }
    return fragment;
  }

  // The countries the enabled kits come from (or the project supplier when no kit is set up yet).
  function guideOrigins() {
    const origins = [...new Set(enabledKitRoutes().map((route) => route.originCountry))];
    return origins.length ? origins : [state.supplier.country];
  }

  function buildGuidesForOrigins(options) {
    const fragment = document.createDocumentFragment();
    guideOrigins().forEach((countryId, index) => {
      const section = document.createElement('div');
      section.className = 'customs-guide-origin';
      if (index) section.classList.add('separated');
      section.append(buildCustomsGuide(countryId, options));
      fragment.append(section);
    });
    return fragment;
  }

  function renderCustomsGuides() {
    const sidebar = el('sidebar-customs-guide');
    if (sidebar?.closest('details')?.open) sidebar.replaceChildren(buildGuidesForOrigins({}));
    if (!dom.customsPopover.hidden) dom.customsPopover.querySelector('.customs-popover-body')?.replaceChildren(buildGuidesForOrigins({ compact: true }));
  }

  function populateOriginSelects() {
    const ids = [...new Set([...Object.keys(ORIGIN_PRESETS), state.supplier.country])].sort((a, b) => originName(a).localeCompare(originName(b)));
    document.querySelectorAll('[data-origin-select]').forEach((select) => {
      populateOriginSelect(select, state.supplier.country, ids);
    });
  }

  function populateOriginSelect(select, selected, providedIds = null) {
    const ids = providedIds || [...new Set([...Object.keys(ORIGIN_PRESETS), selected])].sort((a, b) => originName(a).localeCompare(originName(b)));
    const signature = ids.join(',');
    if (select.dataset.signature !== signature) {
      select.dataset.signature = signature;
      select.replaceChildren(...ids.map((id) => new Option(originName(id), id)));
    }
    if (select !== document.activeElement) select.value = selected;
  }

  function setSupplierOrigin(countryId, { name, lat, lon } = {}) {
    const preset = ORIGIN_PRESETS[countryId];
    state.supplier.country = countryId;
    state.supplier.countryName = preset?.name || name || originName(countryId);
    if (preset) {
      state.supplier.location = preset.place;
      state.supplier.lat = preset.lat;
      state.supplier.lon = preset.lon;
    } else if (Number.isFinite(lat) && Number.isFinite(lon)) {
      state.supplier.location = state.supplier.countryName;
      state.supplier.lat = lat;
      state.supplier.lon = lon;
    }
    populateOriginSelects();
    saveSettingsSoon();
    refreshViews();
  }

  let customsPopoverTimer = 0;
  function bindCustomsPopover() {
    const show = (anchor, delay = 0) => {
      clearTimeout(customsPopoverTimer);
      if (delay) customsPopoverTimer = setTimeout(() => showCustomsPopover(anchor), delay);
      else showCustomsPopover(anchor);
    };
    const scheduleHide = () => {
      if (state.customsPinned) return;
      clearTimeout(customsPopoverTimer);
      customsPopoverTimer = setTimeout(() => { if (!dom.customsPopover.matches(':hover')) hideCustomsPopover(); }, 220);
    };
    document.addEventListener('mouseover', (event) => {
      const anchor = event.target.closest?.('[data-customs-info]');
      if (anchor) {
        if (!dom.customsPopover.hidden && state.customsAnchor === anchor) clearTimeout(customsPopoverTimer);
        else show(anchor, dom.customsPopover.hidden ? 280 : 0);
      } else if (event.target.closest?.('#customs-popover')) {
        clearTimeout(customsPopoverTimer);
      }
    });
    document.addEventListener('mouseout', (event) => {
      if (event.target.closest?.('[data-customs-info], #customs-popover') && !event.relatedTarget?.closest?.('[data-customs-info], #customs-popover')) scheduleHide();
    });
    document.addEventListener('focusin', (event) => {
      const anchor = event.target.closest?.('[data-customs-info]');
      if (anchor) show(anchor);
    });
    document.addEventListener('focusout', (event) => {
      if (event.target.closest?.('[data-customs-info]') && !event.relatedTarget?.closest?.('[data-customs-info], #customs-popover')) scheduleHide();
    });
    document.addEventListener('click', (event) => {
      const pin = event.target.closest?.('[data-customs-pin]');
      if (pin) {
        state.customsPinned = !state.customsPinned;
        pin.setAttribute('aria-expanded', String(state.customsPinned));
        if (state.customsPinned) showCustomsPopover(pin.closest('[data-customs-info]') || pin);
        else hideCustomsPopover();
      } else if (event.target.closest?.('[data-customs-close]')) {
        state.customsPinned = false;
        hideCustomsPopover();
      }
    });
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && !dom.customsPopover.hidden) {
        state.customsPinned = false;
        hideCustomsPopover();
      }
    });
    el('sidebar-customs-guide')?.closest('details')?.addEventListener('toggle', renderCustomsGuides);
  }

  function showCustomsPopover(anchor) {
    const popover = dom.customsPopover;
    state.customsAnchor = anchor;
    const head = document.createElement('div');
    head.className = 'customs-popover-head';
    const title = document.createElement('span');
    title.textContent = `Customs & VAT in ${destinationConfig().name}`;
    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'icon-button';
    close.dataset.customsClose = 'true';
    close.setAttribute('aria-label', 'Close customs rules');
    close.textContent = '×';
    head.append(title, close);
    const body = document.createElement('div');
    body.className = 'customs-popover-body customs-guide';
    body.append(buildGuidesForOrigins({ compact: true }));
    popover.replaceChildren(head, body);
    popover.hidden = false;
    const rect = anchor.getBoundingClientRect();
    const width = Math.min(360, window.innerWidth - 16);
    popover.style.width = `${width}px`;
    const { height } = popover.getBoundingClientRect();
    // Prefer the right of the field, then the left, then below it on narrow screens.
    let left = rect.left;
    let top = rect.bottom + 8;
    if (window.innerWidth - rect.right - 18 >= width) {
      left = rect.right + 10;
      top = rect.top - 20;
    } else if (rect.left - 18 >= width) {
      left = rect.left - width - 10;
      top = rect.top - 20;
    }
    left = clamp(left, 8, window.innerWidth - width - 8);
    top = clamp(top, 8, Math.max(8, window.innerHeight - height - 8));
    popover.style.transform = `translate(${Math.round(left)}px, ${Math.round(top)}px)`;
  }

  function hideCustomsPopover() {
    dom.customsPopover.hidden = true;
    document.querySelectorAll('[data-customs-pin]').forEach((pin) => pin.setAttribute('aria-expanded', 'false'));
  }

  /* ---------- Route map ---------- */

  function setFlowMode(mode) {
    state.flowMode = mode === 'map' ? 'map' : 'route';
    document.querySelectorAll('[data-flow-mode]').forEach((button) => button.setAttribute('aria-pressed', String(button.dataset.flowMode === state.flowMode)));
    dom.flowMap.classList.toggle('hidden', state.flowMode !== 'map');
    dom.flowOverview.querySelector('.flow-scroll').classList.toggle('hidden', state.flowMode === 'map');
    placeFlowBoxes();
    if (state.flowMode === 'map') renderMaps();
  }

  async function ensureMapData() {
    if (state.map.features || state.map.loading) return;
    if (!window.topojson) {
      state.map.error = 'The map library could not be loaded. Check the internet connection and reload.';
      return;
    }
    state.map.loading = true;
    try {
      const response = await fetch(MAP_DATA_URL);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const topology = await response.json();
      const collection = window.topojson.feature(topology, topology.objects.countries);
      state.map.names = new Map();
      state.map.features = collection.features.map((feature) => {
        const id = String(feature.id ?? '').padStart(3, '0');
        const rings = feature.geometry?.type === 'Polygon' ? feature.geometry.coordinates
          : feature.geometry?.type === 'MultiPolygon' ? feature.geometry.coordinates.flat() : [];
        let lonMin = Infinity; let lonMax = -Infinity; let latMin = Infinity; let latMax = -Infinity;
        rings.forEach((ring) => ring.forEach(([lon, lat]) => {
          lonMin = Math.min(lonMin, lon); lonMax = Math.max(lonMax, lon);
          latMin = Math.min(latMin, lat); latMax = Math.max(latMax, lat);
        }));
        const name = feature.properties?.name || id;
        state.map.names.set(id, name);
        return { id, name, rings, bbox: [lonMin, latMin, lonMax, latMax] };
      }).filter((feature) => feature.rings.length && feature.bbox[3] > -60);
      // The map's own English country names get their translations before the map is first drawn.
      I18n?.addPairs('sv', state.map.features.map((feature) => [feature.name, I18n.countryName('sv', feature.id, feature.name)]));
      state.map.error = null;
    } catch (error) {
      console.warn('Map data could not be loaded.', error);
      state.map.error = 'The map outline could not be downloaded. Routes are still listed on the right.';
    } finally {
      state.map.loading = false;
    }
    populateOriginSelects();
    if (state.flowMode === 'map' && state.activeView === VIEW.wire) renderMaps();
  }

  function countryPath(feature, project) {
    return feature.rings.map((ring) => {
      let d = '';
      let previousLon = null;
      ring.forEach(([lon, lat], index) => {
        const [x, y] = project([lon, lat]);
        const jump = previousLon !== null && Math.abs(lon - previousLon) > 180;
        d += `${index === 0 || jump ? 'M' : 'L'}${x.toFixed(1)},${y.toFixed(1)}`;
        previousLon = lon;
      });
      return `${d}Z`;
    }).join('');
  }

  function countryLayer(projection, highlight = {}) {
    const [lonMin, latMin, lonMax, latMax] = projection.bounds;
    return state.map.features
      .filter(({ bbox }) => bbox[2] >= lonMin - 5 && bbox[0] <= lonMax + 5 && bbox[3] >= latMin - 5 && bbox[1] <= latMax + 5)
      .map((feature) => {
        const agreement = originAgreement(feature.id);
        const classes = ['map-country', agreement];
        if (feature.id === highlight.supplier) classes.push('supplier');
        if (feature.id === highlight.compare) classes.push('compare');
        return `<path class="${classes.join(' ')}" d="${countryPath(feature, projection.project)}" data-country-id="${feature.id}"><title>${escapeMarkup(`${feature.name} · ${AGREEMENT_TEXT[agreement].badge}`)}</title></path>`;
      }).join('');
  }

  // A stretch of a curve as an SVG path, sampled so a box can hide the middle of a wire.
  function curvePath(curve, from = 0, to = 1, steps = 36) {
    let d = '';
    for (let index = 0; index <= steps; index += 1) {
      const [x, y] = curve.at(from + ((to - from) * index) / steps);
      d += `${index ? 'L' : 'M'}${x.toFixed(1)},${y.toFixed(1)}`;
    }
    return d;
  }

  function mapArrowhead(curve, t, className, style = '') {
    const [x, y] = curve.at(t);
    return `<path class="map-arrow ${className}"${style} transform="translate(${x.toFixed(1)} ${y.toFixed(1)}) rotate(${curve.angle(t).toFixed(1)})" d="M-11,-5 L0,0 L-11,5 Z"/>`;
  }

  // A wire with a wide invisible hit line (hover and click open the info panel) and an arrowhead at its end.
  function mapWire(curve, { from = 0, to = 1, target = '', label = '', className = '', color = '' } = {}) {
    const d = curvePath(curve, from, to);
    const hit = target
      ? `<path class="map-route-hit" d="${d}" data-map-target="${escapeMarkup(target)}" tabindex="0" role="button" aria-label="${escapeMarkup(label)}"/>`
      : '';
    return `${hit}<path class="map-route ${className}"${color ? ` style="stroke:${color}"` : ''} d="${d}"/>${mapArrowhead(curve, to, className, color ? ` style="fill:${color}"` : '')}`;
  }

  function mapPoint([x, y], target, label, kind, anchor = 'start') {
    const below = anchor === 'below';
    const textAnchor = below ? 'middle' : anchor;
    const dx = textAnchor === 'start' ? 9 : textAnchor === 'end' ? -9 : 0;
    const dy = below ? 20 : textAnchor === 'middle' ? -10 : 4;
    return `<g class="map-point ${kind}" data-map-target="${target}" tabindex="0" role="button" aria-label="${escapeMarkup(label)}"><circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="12" class="map-hit-circle"/><circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="5.5"/><text x="${(x + dx).toFixed(1)}" y="${(y + dy).toFixed(1)}" text-anchor="${textAnchor}">${escapeMarkup(label)}</text></g>`;
  }

  function renderMaps() {
    if (state.flowMode !== 'map') return;
    ensureMapData();
    renderMapScene();
    renderMapInfo();
  }

  // Kits from the same manufacturer and country share one route, e.g. four TEI kits from the USA.
  function enabledKitRoutes() {
    if (!state.workbook) return [];
    const routes = new Map();
    state.workbook.SheetNames.filter((name) => state.mappings[name]?.enabled && isSourceKit(name)).forEach((sheetName) => {
      const kit = kitSettings(sheetName);
      const manufacturer = kit.manufacturer || sheetName;
      const key = `${manufacturer}|${kit.originCountry}`;
      if (!routes.has(key)) {
        routes.set(key, {
          key,
          manufacturer,
          originCountry: kit.originCountry,
          origin: ORIGIN_PRESETS[kit.originCountry] || ORIGIN_PRESETS[SUPPLIER_DEFAULTS.country],
          sheetNames: [],
          kits: [],
          color: GROUP_COLORS[routes.size % GROUP_COLORS.length],
        });
      }
      const route = routes.get(key);
      route.sheetNames.push(sheetName);
      route.kits.push(kit);
    });
    return [...routes.values()];
  }

  // The routes the map draws: one per manufacturer and country, or the project supplier before a workbook is open.
  function mapRoutes() {
    const routes = enabledKitRoutes();
    return routes.length ? routes : [{
      key: 'supplier',
      manufacturer: state.supplier.name,
      originCountry: state.supplier.country,
      origin: { lon: state.supplier.lon, lat: state.supplier.lat, name: originName() },
      sheetNames: [],
      kits: [],
      color: GROUP_COLORS[0],
    }];
  }

  // Where each map panel sits inside the stage, in pixels.
  function mapPanels() {
    // Layout pixels (offset*, client*) rather than getBoundingClientRect, so CSS transforms and page zoom do not skew the geometry.
    const measure = (canvas) => {
      const panel = canvas.parentElement;
      return { x: panel.offsetLeft + panel.clientLeft, y: panel.offsetTop + panel.clientTop, w: canvas.clientWidth, h: canvas.clientHeight };
    };
    return { world: measure(dom.mapWorld), nordic: measure(dom.mapNorway) };
  }

  const boxSize = (box) => ({ w: box.offsetWidth, h: box.offsetHeight });

  // A coarse picture of where a country's land is on a panel, so a table can be kept off it.
  function landMask(view, width, height, countryIds) {
    const scale = 0.25;
    const columns = Math.max(1, Math.ceil(width * scale));
    const rows = Math.max(1, Math.ceil(height * scale));
    const canvas = document.createElement('canvas');
    canvas.width = columns;
    canvas.height = rows;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (!context || typeof Path2D === 'undefined') return null;
    context.scale(scale, scale);
    state.map.features.filter((feature) => countryIds.includes(feature.id)).forEach((feature) => context.fill(new Path2D(countryPath(feature, view.project))));
    const pixels = context.getImageData(0, 0, columns, rows).data;
    return (rect) => {
      const x0 = Math.max(0, Math.floor(rect.x * scale));
      const x1 = Math.min(columns - 1, Math.floor((rect.x + rect.w) * scale));
      const y0 = Math.max(0, Math.floor(rect.y * scale));
      const y1 = Math.min(rows - 1, Math.floor((rect.y + rect.h) * scale));
      for (let y = y0; y <= y1; y += 1) {
        for (let x = x0; x <= x1; x += 1) if (pixels[(y * columns + x) * 4 + 3] > 40) return true;
      }
      return false;
    };
  }

  // The map is two panels: the ocean crossing on the left and the destination country enlarged on the right. The
  // shipment table sits on the ocean arrow and the customer table on the arrow north to the customer; every wire
  // is drawn up to a table, hidden behind it, and drawn on from its far side.
  function renderMapScene() {
    const stage = dom.mapStage;
    const width = stage.clientWidth;
    if (!width) return;
    const stacked = width < MAP_OVERLAY_MIN_WIDTH;
    const ready = Boolean(state.map.features);
    const overlay = ready && !stacked;
    stage.classList.toggle('stacked', stacked);
    stage.querySelector('.map-panels').style.setProperty('--map-height', `${Math.max(540, Math.min(640, Math.round(width * 0.56)))}px`);
    dom.mapBoxes.classList.toggle('flow', !overlay);
    [...[...state.laneRefs.entries()].map(([maker, refs]) => [`ship:${maker}`, refs.box]), ['warehouse', dom.flowBoxWarehouse], ['customer', dom.flowBoxCustomer]].forEach(([name, box]) => {
      box.classList.toggle('is-collapsed', !state.map.open[name]);
      setBoxToggle(box, state.map.open[name]);
    });
    if (!ready) {
      state.map.worldKey = null;
      state.map.nordicKey = null;
      dom.mapWorld.innerHTML = `<p class="map-status">${escapeMarkup(state.map.error || 'Loading map outlines…')}</p>`;
      dom.mapNorway.innerHTML = '';
      dom.mapOverlay.innerHTML = '';
      dom.mapLegend.innerHTML = '';
      resetFlowBoxes();
      return;
    }

    const { world, nordic } = mapPanels();
    const destination = destinationConfig();
    const customer = customerConfig();
    const routes = mapRoutes();
    const compareFeature = state.map.compare ? state.map.features.find((feature) => feature.id === state.map.compare) : null;
    const compare = compareFeature ? compareFeature.centroid || (compareFeature.centroid = featureCentroid(compareFeature)) : null;

    const dist = ([x1, y1], [x2, y2]) => Math.hypot(x2 - x1, y2 - y1);
    const around = (point, stub, extra = 0) => Core.rectAround(point, { w: 2 * stub + extra, h: 2 * stub });
    // Keeps a full stretch of wire visible on each side of a table; when the panel is too crowded for that, settles for less.
    // A table goes where it covers no land of the destination country if there is room, otherwise anywhere clear.
    const place = (size, anchor, bounds, avoidFor, blocked = null) => {
      let result = null;
      for (const useBlocked of blocked ? [true, false] : [false]) {
        for (const stub of [MAP_STUB, 20, 10]) {
          result = Core.placeBox(size, anchor, { bounds, avoid: avoidFor(stub), blocked: useBlocked ? blocked : null, step: 8 });
          if (result.ok) return result;
        }
      }
      return result;
    };
    const centreOf = (rect) => [rect.x + rect.w / 2, rect.y + rect.h / 2];
    const viaBox = (from, rect, to, bow) => {
      const centre = centreOf(rect);
      const into = Core.quadraticBow(from, centre, bow);
      const out = Core.quadraticBow(centre, to, bow);
      return { into, out, enter: Core.curveAroundRect(into.at, rect)?.enter ?? 1, exit: Core.curveAroundRect(out.at, rect)?.exit ?? 0 };
    };

    // ---- Left panel: from the manufacturers to the destination country.
    const spread = [...routes.map(({ origin }) => [origin.lon, origin.lat]), ...(compare ? [compare] : [])];
    const lons = spread.map(([lon]) => lon);
    const lats = spread.map(([, lat]) => lat);
    let bounds = [
      Math.min(...lons, destination.bounds[0]) - 16,
      Math.min(...lats, destination.bounds[1]) - 10,
      Math.max(...lons, destination.bounds[2]) + 6,
      Math.max(...lats, destination.bounds[3]) + 4,
    ];
    if (bounds[2] - bounds[0] < 60) {
      const middle = (bounds[0] + bounds[2]) / 2;
      bounds = [middle - 30, bounds[1], middle + 30, bounds[3]];
    }
    let worldView = Core.fitProjection(bounds, world.w, world.h);
    const overshoot = worldView.bounds[3] - 84;       // no empty polar sea above Greenland
    if (overshoot > 0) worldView = Core.fitProjection([bounds[0], bounds[1] - overshoot, bounds[2], bounds[3] - overshoot], world.w, world.h);
    const worldKey = [world.w, world.h, worldView.bounds.map((value) => Math.round(value * 10)).join(','), routes[0].originCountry, state.map.compare || ''].join('|');
    if (state.map.worldKey !== worldKey) {
      state.map.worldKey = worldKey;
      dom.mapWorld.innerHTML = `<svg viewBox="0 0 ${world.w} ${world.h}" role="img" aria-label="Supply routes to ${escapeMarkup(destination.name)}"><rect class="map-sea" width="${world.w}" height="${world.h}"/><g class="map-countries">${countryLayer(worldView, { supplier: routes[0].originCountry, compare: state.map.compare })}</g></svg>`;
    }

    const arrivalWorld = worldView.project(destination.arrival);
    const legs = routes.map((route) => ({ route, from: worldView.project([route.origin.lon, route.origin.lat]), box: state.laneRefs.get(route.manufacturer)?.box || null, rect: null }));
    // The enlarged country is outlined on this panel; no table may cover it.
    const [frameWest, frameNorth] = worldView.project([destination.bounds[0], destination.bounds[3]]);
    const [frameEast, frameSouth] = worldView.project([destination.bounds[2], destination.bounds[1]]);
    const frameRect = { x: frameWest - 8, y: frameNorth - 8, w: frameEast - frameWest + 16, h: frameSouth - frameNorth + 16 };
    if (overlay) {
      // Each manufacturer's table sits on its own arrow, the farthest route first; later tables avoid earlier ones.
      const placed = [];
      [...legs].sort((a, b) => dist(b.from, arrivalWorld) - dist(a.from, arrivalWorld)).forEach((leg) => {
        if (!leg.box) return;
        leg.rect = place(boxSize(leg.box), Core.quadraticBow(leg.from, arrivalWorld, 0.16).at(0.5), { x: 6, y: 34, w: world.w - 12, h: world.h - 40 }, (stub) => [
          ...legs.map(({ from }) => around(from, stub, 120)),
          around(arrivalWorld, stub, 60),
          frameRect,
          ...placed,
          ...(compare ? [around(worldView.project(compare), stub, 100)] : []),
        ]);
        placed.push(leg.rect);
      });
    }
    const worldParts = [];
    if (overlay) {
      worldParts.push(`<rect class="map-zoom-frame" x="${frameWest.toFixed(1)}" y="${frameNorth.toFixed(1)}" width="${(frameEast - frameWest).toFixed(1)}" height="${(frameSouth - frameNorth).toFixed(1)}" rx="3"/>`);
    }
    if (compare) {
      const from = worldView.project(compare);
      worldParts.push(mapWire(Core.quadraticBow(from, arrivalWorld, 0.16), { target: 'compare', label: 'Comparison line', className: 'compare' }));
    }
    legs.forEach(({ route, from, rect }, index) => {
      const target = `kit:${route.key}`;
      const label = `${route.manufacturer}: ${originName(route.originCountry)} to ${destination.name}`;
      const bow = 0.15 + (index % 3) * 0.04;
      if (rect) {
        const via = viaBox(from, rect, arrivalWorld, bow);
        worldParts.push(mapWire(via.into, { to: via.enter, target, label, color: route.color }));
        worldParts.push(mapWire(via.out, { from: via.exit, target, label, className: 'trunk', color: route.color }));
      } else {
        worldParts.push(mapWire(Core.quadraticBow(from, arrivalWorld, bow), { target, label, color: route.color }));
      }
    });
    if (compare) worldParts.push(mapPoint(worldView.project(compare), 'compare', originName(state.map.compare), 'compare', 'below'));
    legs.forEach(({ route, from }) => worldParts.push(mapPoint(from, `kit:${route.key}`, route.manufacturer, 'supplier', 'below')));
    worldParts.push(mapPoint(arrivalWorld, 'customs', `${destination.name} customs`, 'customs', 'below'));

    // ---- Right panel: the destination country, enlarged.
    const extent = isCrossBorderSale()
      ? [Math.min(destination.bounds[0], customer.bounds[0]), Math.min(destination.bounds[1], customer.bounds[1]), Math.max(destination.bounds[2], customer.bounds[2]), Math.max(destination.bounds[3], customer.bounds[3])]
      : destination.bounds;
    const nordicView = Core.fitProjection(extent, nordic.w, nordic.h);
    const nordicKey = [nordic.w, nordic.h, state.warehouse.country, state.customer.country].join('|');
    if (state.map.nordicKey !== nordicKey) {
      state.map.nordicKey = nordicKey;
      dom.mapNorway.innerHTML = `<svg viewBox="0 0 ${nordic.w} ${nordic.h}" role="img" aria-label="Map of ${escapeMarkup(destination.name)} with customs, warehouse and delivery routes"><rect class="map-sea" width="${nordic.w}" height="${nordic.h}"/><g class="map-countries">${countryLayer(nordicView)}</g></svg>`;
    }
    const cityPoint = ([, lat, lon]) => nordicView.project([lon, lat]);
    const landIds = [...new Set([state.warehouse.country, state.customer.country])];
    const landKey = `${nordicKey}|${landIds.join(',')}`;
    if (state.map.landKey !== landKey) {
      state.map.landKey = landKey;
      state.map.land = landMask(nordicView, nordic.w, nordic.h, landIds);
    }
    const north = customer.customers.reduce((best, city) => (city[1] > best[1] ? city : best), customer.customers[0]);
    const customs = nordicView.project(destination.arrival);
    const warehouse = nordicView.project(destination.warehouse);
    const end = cityPoint(north);
    let customerRect = null;
    let warehouseRect = null;
    if (overlay) {
      const nordicBounds = { x: 6, y: 34, w: nordic.w - 12, h: nordic.h - 40 };
      const endpoints = (stub) => [around(customs, stub, 90), around(warehouse, stub, 110), around(end, stub, 90)];
      customerRect = place(boxSize(dom.flowBoxCustomer), [(warehouse[0] + end[0]) / 2, (warehouse[1] + end[1]) / 2], nordicBounds, endpoints, state.map.land);
      const size = boxSize(dom.flowBoxWarehouse);
      warehouseRect = place(size, [warehouse[0] + 70 + size.w / 2, warehouse[1]], nordicBounds, (stub) => [...endpoints(stub), customerRect], state.map.land);
    }
    const nordicParts = [];
    customer.customers.forEach((city) => {
      if (city === north) return;
      const [x, y] = cityPoint(city);
      // A city that shares its spot with the customs or warehouse marker is already labelled by that marker.
      if ([customs, warehouse].some((point) => Math.hypot(point[0] - x, point[1] - y) < 14)) return;
      nordicParts.push(`<g class="map-city" data-map-target="delivery"><circle class="map-hit-circle" cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="10"/><circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3"/><text x="${(x - 6).toFixed(1)}" y="${(y + 3.5).toFixed(1)}" text-anchor="end">${escapeMarkup(city[0])}</text></g>`);
    });
    const receive = Core.quadraticBow(customs, warehouse, 0.35);
    nordicParts.push(`<path class="map-route-hit" d="${curvePath(receive)}" data-map-target="warehouse"/><path class="map-route receive" d="${curvePath(receive)}"/>`);
    if (customerRect) {
      const via = viaBox(warehouse, customerRect, end, 0.1);
      nordicParts.push(mapWire(via.into, { to: via.enter, target: 'delivery', label: 'Delivery', className: 'delivery trunk' }));
      nordicParts.push(mapWire(via.out, { from: via.exit, target: 'delivery', label: 'Delivery', className: 'delivery trunk' }));
    } else {
      nordicParts.push(mapWire(Core.quadraticBow(warehouse, end, 0.1), { target: 'delivery', label: 'Delivery', className: 'delivery trunk' }));
    }
    if (warehouseRect) {
      const [x, y] = Core.rayToRect(centreOf(warehouseRect), warehouse, warehouseRect);
      nordicParts.push(`<path class="map-link" d="M${warehouse[0].toFixed(1)},${warehouse[1].toFixed(1)}L${x.toFixed(1)},${y.toFixed(1)}"/>`);
    }
    nordicParts.push(mapPoint(customs, 'customs', 'Customs', 'customs', 'end'));
    nordicParts.push(mapPoint(warehouse, 'warehouse', state.warehouse.site || 'Warehouse', 'warehouse', 'below'));
    nordicParts.push(mapPoint(end, 'delivery', north[0], 'customer', end[0] > nordic.w - 100 ? 'end' : 'start'));

    // ---- One overlay carries every wire and point; the dashed wire leads from the ocean panel into the enlargement.
    let zoom = '';
    if (overlay) {
      const from = [world.x + arrivalWorld[0], world.y + arrivalWorld[1]];
      const to = [nordic.x + customs[0], nordic.y + customs[1]];
      zoom = mapWire(Core.quadraticBow(from, to, 0.08), { to: 0.93, className: 'zoom' });
    }
    const extentHeight = Math.max(world.y + world.h, nordic.y + nordic.h);
    dom.mapOverlay.setAttribute('width', String(width));
    dom.mapOverlay.setAttribute('height', String(extentHeight));
    dom.mapOverlay.setAttribute('viewBox', `0 0 ${width} ${extentHeight}`);
    dom.mapOverlay.innerHTML = `<g transform="translate(${world.x.toFixed(1)} ${world.y.toFixed(1)})">${worldParts.join('')}</g><g transform="translate(${nordic.x.toFixed(1)} ${nordic.y.toFixed(1)})">${nordicParts.join('')}</g>${zoom}`;

    const put = (box, rect, panel) => {
      box.style.left = `${Math.round(panel.x + rect.x)}px`;
      box.style.top = `${Math.round(panel.y + rect.y)}px`;
    };
    if (overlay) {
      legs.forEach(({ box, rect }) => { if (box && rect) put(box, rect, world); });
      put(dom.flowBoxCustomer, customerRect, nordic);
      put(dom.flowBoxWarehouse, warehouseRect, nordic);
    } else {
      resetFlowBoxes();
    }
    dom.mapLegend.innerHTML = `${routes.map(({ manufacturer, originCountry, color }) => `<span><i style="background:${color}"></i>${escapeMarkup(`${manufacturer} · ${originName(originCountry)}`)}</span>`).join('')}<span><i class="eu"></i>EU / EEA</span><span><i class="efta"></i>EFTA</span><span><i class="usa"></i>USA</span>`;
  }

  function featureCentroid(feature) {
    const largest = feature.rings.reduce((best, ring) => (ring.length > best.length ? ring : best), []);
    const sum = largest.reduce(([sx, sy], [lon, lat]) => [sx + lon, sy + lat], [0, 0]);
    return [sum[0] / largest.length, sum[1] / largest.length];
  }

  function renderMapInfo() {
    const target = state.map.hover || state.map.pinned || 'overview';
    const destination = destinationConfig();
    const model = calculateProfitability();
    const items = state.consolidated;
    const sum = (field) => items.reduce((total, item) => total + (item[field] ?? 0), 0);
    const nodes = [];
    const heading = (eyebrow, title) => {
      const p = document.createElement('p');
      p.className = 'eyebrow';
      p.textContent = eyebrow;
      const h = document.createElement('h4');
      h.textContent = title;
      nodes.push(p, h);
    };
    const stats = (pairs) => {
      const dl = document.createElement('dl');
      dl.className = 'stat-list';
      setStats(dl, pairs);
      nodes.push(dl);
    };
    const note = (text) => {
      const p = document.createElement('p');
      p.className = 'fine-print';
      p.textContent = text;
      nodes.push(p);
    };
    const guide = (countryId) => {
      const wrap = document.createElement('div');
      wrap.className = 'customs-guide';
      wrap.append(buildCustomsGuide(countryId, { compact: true }));
      nodes.push(wrap);
    };

    if (typeof target === 'string' && target.startsWith('kit:')) {
      const route = enabledKitRoutes().find((candidate) => candidate.key === target.slice(4));
      if (route) {
        const unique = (values) => [...new Set(values)].join(', ');
        const lines = state.consolidated.filter((item) => route.sheetNames.includes(item.pricingSource));
        heading('Kit route', `${route.manufacturer} · ${originName(route.originCountry)} → ${destination.name}`);
        stats([
          [route.sheetNames.length === 1 ? 'Source sheet' : 'Source sheets', route.sheetNames.join(', ')],
          ['Customs treatment', TREATMENT_LABELS[Core.customsTreatment(route.originCountry, state.warehouse.country)]],
          ['Source currency', unique(route.kits.map(kitCurrency))],
          ['Discount', unique(route.kits.map((kit) => formatPercent(kitDiscount(kit))))],
          ['Customs duty', unique(lines.map((item) => formatPercent(item.dutyRate ?? 0)))],
          ['Freight currency', unique(route.kits.map((kit) => kit.shippingCurrency))],
          ['Purchase after discount', formatWhole(lines.reduce((total, item) => total + (item.discountedTotal ?? 0), 0))],
          ['Freight & import total', formatWhole(lines.reduce((total, item) => total + (item.freightAndImport ?? 0), 0))],
        ]);
        guide(route.originCountry);
      }
    } else if (target === 'supplier') {
      heading('Manufacturer', state.supplier.name);
      stats([['Location', `${state.supplier.location} · ${originName()}`], ['Discount', formatPercent(state.config.discount)], ['List value', formatWhole(sum('convertedTotal'))], ['Purchase after discount', formatWhole(sum('discountedTotal'))]]);
    } else if (target === 'route') {
      heading('Freight & tolls', `${[...new Set(guideOrigins().map((countryId) => originName(countryId)))].join(', ')} → ${destination.name}`);
      stats([
        ['Consolidated shipment', formatWhole(sum('consolidatedFreight'))],
        ['Kit freight', formatWhole(sum('kitFreight'))],
        ['Customs duty', `${formatPercent(state.freight.dutyRate)} · ${formatWhole(items.reduce((total, item) => total + (item.duty ?? 0), 0))}`],
        ['Insurance & clearance', formatWhole(items.reduce((total, item) => total + (item.insurance ?? 0) + (item.clearance ?? 0), 0))],
        ['Freight & import total', formatWhole(sum('freightAndImport'))],
      ]);
      guideOrigins().forEach((countryId) => guide(countryId));
    } else if (target === 'compare' && state.map.compare) {
      heading('Comparison line', `${originName(state.map.compare)} → ${destination.name}`);
      guide(state.map.compare);
      const actions = document.createElement('div');
      actions.className = 'map-actions';
      const use = document.createElement('button');
      use.type = 'button';
      use.className = 'button small primary';
      use.dataset.mapUseOrigin = state.map.compare;
      use.textContent = `Use ${originName(state.map.compare)} as supplier country`;
      const clear = document.createElement('button');
      clear.type = 'button';
      clear.className = 'button small ghost';
      clear.dataset.mapClearCompare = 'true';
      clear.textContent = 'Remove line';
      actions.append(use, clear);
      nodes.push(actions);
    } else if (target === 'customs') {
      heading(`${destination.name} customs`, 'Import clearance');
      stats([['Landed value', formatWhole(sum('landedCost'))], ['Import VAT on this order', formatWhole(sum('importVat'))], ['Duty', formatWhole(items.reduce((total, item) => total + (item.duty ?? 0), 0))]]);
      guideOrigins().forEach((countryId) => guide(countryId));
    } else if (target === 'warehouse') {
      heading('Warehouse', state.warehouse.site || 'Warehouse');
      stats([
        ['Locations', `${model.warehouse.shelfLocations} shelf · ${model.warehouse.drawerLocations} drawer · ${model.warehouse.pallets} pallet`],
        ['Per month', `${formatNok(model.warehouse.monthlyNok)} · ${formatWhole(model.warehouse.monthlyTotal)}`],
        ['Storage / unit / month', formatWhole(model.perUnit.storagePerUnitMonth)],
      ]);
    } else if (target === 'delivery') {
      heading('Delivery', `Warehouse → customers in ${customerConfig().name}`);
      stats([
        ['Parcel ≤35 kg (business)', formatNok(state.warehouse.rates.parcel)],
        ['Private recipient surcharge', formatNok(state.warehouse.rates.privateSurcharge)],
        ['Shipment / unit sold', formatWhole(model.perUnit.outboundPerUnit)],
        ['Outtake / unit sold', formatWhole(model.perUnit.outtakePerUnit)],
        ['Expected sales / yr', `${formatNumber(model.salesUnits, 0)} units · ${formatWhole(model.revenue)}`],
      ]);
      note(state.sales.customerPaysOutbound ? 'Outbound freight is recharged to the customer.' : 'Outbound freight is counted as a cost in the dashboard.');
    } else if (typeof target === 'string' && target.startsWith('country:')) {
      const id = target.slice(8);
      heading('Country', originName(id));
      note(`${AGREEMENT_TEXT[originAgreement(id)].badge}. Click to draw a comparison line to ${destination.name}.`);
    } else {
      heading('Route overview', `${enabledKitRoutes().length} kit route${enabledKitRoutes().length === 1 ? '' : 's'} → ${destination.name}`);
      stats([
        ['Purchase after discount', formatWhole(sum('discountedTotal'))],
        ['Freight & import', formatWhole(sum('freightAndImport'))],
        [`Landed in ${destination.name}`, formatWhole(sum('landedCost'))],
        ['Warehouse / month', formatWhole(model.warehouse.monthlyTotal)],
        ['Net profit / yr', formatWhole(model.net)],
      ]);
      note('Hover a wire or point for details; click one to keep it here. Click any country to compare its customs rules.');
    }
    if (state.map.pinned && !state.map.hover) {
      const unpin = document.createElement('button');
      unpin.type = 'button';
      unpin.className = 'link-button';
      unpin.dataset.mapUnpin = 'true';
      unpin.textContent = 'Back to overview';
      nodes.push(unpin);
    }
    dom.mapInfo.replaceChildren(...nodes);
  }

  function bindMapEvents() {
    const targetOf = (event) => {
      const element = event.target.closest?.('[data-map-target]');
      if (element) return element.dataset.mapTarget;
      const country = event.target.closest?.('[data-country-id]');
      return country ? `country:${country.dataset.countryId}` : null;
    };
    [dom.mapStage].forEach((container) => {
      container.addEventListener('mouseover', (event) => {
        const target = targetOf(event);
        if (target !== state.map.hover) {
          state.map.hover = target;
          renderMapInfo();
        }
      });
      container.addEventListener('mouseleave', () => {
        state.map.hover = null;
        renderMapInfo();
      });
      container.addEventListener('focusin', (event) => {
        state.map.hover = targetOf(event);
        renderMapInfo();
      });
      container.addEventListener('focusout', () => {
        state.map.hover = null;
        renderMapInfo();
      });
      container.addEventListener('click', (event) => {
        const target = targetOf(event);
        if (!target) return;
        if (target.startsWith('country:')) {
          const id = target.slice(8);
          if (id === state.warehouse.country) return;
          state.map.compare = id === state.supplier.country ? null : id;
          state.map.pinned = state.map.compare ? 'compare' : 'route';
          state.map.hover = null;
          renderMaps();
        } else {
          state.map.pinned = target;
          state.map.hover = null;
        }
        renderMapInfo();
      });
      container.addEventListener('keydown', (event) => {
        if ((event.key === 'Enter' || event.key === ' ') && event.target.closest?.('[data-map-target]')) {
          event.preventDefault();
          state.map.pinned = event.target.closest('[data-map-target]').dataset.mapTarget;
          renderMapInfo();
        }
      });
    });
    dom.mapInfo.addEventListener('click', (event) => {
      const use = event.target.closest('[data-map-use-origin]');
      if (use) {
        const feature = state.map.features?.find((item) => item.id === use.dataset.mapUseOrigin);
        const [lon, lat] = feature ? feature.centroid || featureCentroid(feature) : [];
        state.map.compare = null;
        state.map.pinned = 'route';
        setSupplierOrigin(use.dataset.mapUseOrigin, { name: feature?.name, lat, lon });
        showToast(`Supplier country set to ${originName()}. Customs guidance updated.`);
      }
      if (event.target.closest('[data-map-clear-compare]')) {
        state.map.compare = null;
        state.map.pinned = null;
        renderMaps();
      }
      if (event.target.closest('[data-map-unpin]')) {
        state.map.pinned = null;
        renderMapInfo();
      }
    });
    document.addEventListener('click', (event) => {
      const toggle = event.target.closest?.('[data-map-box-toggle]');
      if (!toggle) return;
      const name = toggle.dataset.mapBoxToggle;
      state.map.open[name] = !state.map.open[name];
      renderFlowLanes();
      renderMaps();
    });
    document.querySelectorAll('[data-flow-mode]').forEach((button) => button.addEventListener('click', () => setFlowMode(button.dataset.flowMode)));
  }

  function summaryCard(label, value) {
    const card = document.createElement('article');
    card.className = 'summary-card';
    const title = document.createElement('span');
    title.textContent = label;
    const number = document.createElement('strong');
    number.textContent = value;
    card.append(title, number);
    return card;
  }

  // A summary card that jumps somewhere useful. Rendered as a real <button> so it is reachable by
  // keyboard and announced as an action rather than as a number.
  function actionCard(label, value, onActivate, options = {}) {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'summary-card is-actionable' + (options.attention ? ' is-attention' : '');
    const title = document.createElement('span');
    title.textContent = label;
    const number = document.createElement('strong');
    number.textContent = value;
    card.append(title, number);
    card.title = options.hint || label;
    card.setAttribute('aria-label', label + ': ' + value + '. ' + (options.hint || ''));
    card.addEventListener('click', onActivate);
    return card;
  }

  async function refreshRates(force) {
    if (!state.workbook) return;
    const target = state.config.outputCurrency;
    const sources = [...new Set(state.consolidated.map((item) => item.currency).filter(Boolean))];
    const kitCurrencies = state.workbook.SheetNames
      .filter((name) => state.mappings[name]?.enabled)
      .flatMap((name) => [kitSettings(name).shippingCurrency, kitCurrency(kitSettings(name))]);
    const requestedPairs = uniqueBy([
      ...sources.map((source) => ({ from: source, to: target })),
      { from: 'USD', to: target },
      { from: 'NOK', to: target },
      ...kitCurrencies.map((currency) => ({ from: currency, to: target })),
      { from: state.freight.consolidatedCurrency, to: target },
      { from: state.freight.clearanceCurrency, to: target },
      ...state.liveRateCells.map(({ from, to }) => ({ from, to })),
    ], (pair) => `${pair.from}/${pair.to}`);
    if (!requestedPairs.length) {
      dom.rateStatus.textContent = 'No currencies were found in the enabled parts sheets.';
      return;
    }

    dom.refreshRates.disabled = true;
    dom.rateStatus.textContent = `Loading rates to ${target}…`;
    const errors = [];

    await Promise.all(requestedPairs.map(async ({ from, to }) => {
      const pairKey = `${from}/${to}`;
      if (from === to) {
        state.pairRates[pairKey] = 1;
        if (to === target) {
          state.rates[from] = 1;
          state.rateDates[from] = new Date().toISOString().slice(0, 10);
        }
        return;
      }
      const cacheKey = `partslist-rate-${from}-${to}`;
      const cached = readRateCache(cacheKey);
      if (!force && cached && Date.now() - cached.savedAt < RATE_CACHE_MS) {
        state.pairRates[pairKey] = cached.rate;
        if (to === target) {
          state.rates[from] = cached.rate;
          state.rateDates[from] = cached.date;
        }
        return;
      }

      try {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 10000);
        const response = await fetch(`${RATE_SOURCE}/rate/${from.toLowerCase()}/${to.toLowerCase()}`, { signal: controller.signal });
        clearTimeout(timeout);
        if (!response.ok) throw new Error(`${pairKey}: HTTP ${response.status}`);
        const data = await response.json();
        const rate = toNumber(data.rate);
        if (rate === null || rate <= 0) throw new Error(`${pairKey}: invalid rate`);
        const rateDate = data.date || new Date().toISOString().slice(0, 10);
        state.pairRates[pairKey] = rate;
        if (to === target) {
          state.rates[from] = rate;
          state.rateDates[from] = rateDate;
        }
        try {
          localStorage.setItem(cacheKey, JSON.stringify({ rate, date: rateDate, savedAt: Date.now() }));
        } catch { /* Cache is optional. */ }
      } catch (error) {
        if (cached) {
          state.pairRates[pairKey] = cached.rate;
          if (to === target) {
            state.rates[from] = cached.rate;
            state.rateDates[from] = cached.date;
          }
          errors.push(`${pairKey} (cached)`);
        } else {
          delete state.pairRates[pairKey];
          if (to === target) delete state.rates[from];
          errors.push(pairKey);
        }
      }
    }));

    dom.refreshRates.disabled = false;
    applyLiveRatesToEngine();
    rebuildConsolidation();
    if (force) markProjectDirty();
    const dates = [...new Set(Object.values(state.rateDates).filter(Boolean))].sort();
    dom.rateStatus.textContent = errors.length
      ? `Some rates could not be refreshed: ${errors.join(', ')}. Missing conversions remain blank.`
      : `Daily reference rates loaded${dates.length ? ` for ${dates[dates.length - 1]}` : ''}. Source: Frankfurter.`;
  }

  function applyLiveRatesToEngine() {
    state.liveRateCells.forEach((cell) => {
      const original = state.matrices[cell.sheetName]?.[cell.row]?.[cell.col];
      if (!extractCurrencyPair(original)) return;
      const rate = state.pairRates[`${cell.from}/${cell.to}`];
      if (!Number.isFinite(rate)) return;
      state.hf.setCellContents({ sheet: state.sheetIds[cell.sheetName], row: cell.row, col: cell.col }, [[rate]]);
    });
  }

  function readRateCache(key) {
    try {
      const value = JSON.parse(localStorage.getItem(key));
      if (value && Number.isFinite(value.rate) && Number.isFinite(value.savedAt)) return value;
    } catch { /* Ignore unavailable or invalid cache data. */ }
    return null;
  }

  function renderTabs() {
    dom.tabs.replaceChildren(
      createTab('Consolidated', VIEW.consolidated),
      createTab('Dashboard', VIEW.dashboard),
      createTab('Warehouse', VIEW.warehouse),
      createTab('Wire view', VIEW.wire),
      // Sheets that PartsList itself generated (Consolidated, Profitability, settings…) are not source sheets.
      ...state.workbook.SheetNames.filter((sheetName) => !isGeneratedSheet(state.workbook.Sheets[sheetName])).map((sheetName) => createTab(sheetName, sheetName)),
    );
  }

  function createTab(label, view) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `view-tab${state.activeView === view ? ' active' : ''}${SPECIAL_VIEWS.has(view) ? ' app-tab' : ''}`;
    button.dataset.view = view;
    button.textContent = label;
    button.addEventListener('click', () => showView(view));
    return button;
  }

  // The current view lives in the URL hash (with '#' + name), so views can be bookmarked and linked to,
  // and the browser Back button steps through them. Sheet names keep their own '#' form.
  function viewFromHash() {
    const raw = decodeURIComponent(String(location.hash || '').replace(/^#/, ''));
    if (!raw) return null;
    if (Object.values(VIEW).includes(raw)) return raw;
    return state.workbook && state.workbook.SheetNames.includes(raw) ? raw : null;
  }

  function syncHash(view) {
    const next = '#' + encodeURIComponent(view);
    if (location.hash !== next) {
      state.suppressHash = true;
      history.pushState(null, '', next);
      state.suppressHash = false;
    }
  }

  function applyHashView() {
    const view = viewFromHash();
    if (!view || view === state.activeView) return;
    showView(view);
  }

  function showView(view) {
    state.activeView = view;
    const isSheet = !SPECIAL_VIEWS.has(view);
    dom.welcome.classList.add('hidden');
    dom.consolidatedView.classList.toggle('hidden', view !== VIEW.consolidated);
    dom.dashboardView.classList.toggle('hidden', view !== VIEW.dashboard);
    dom.warehouseView.classList.toggle('hidden', view !== VIEW.warehouse);
    dom.wireView.classList.toggle('hidden', view !== VIEW.wire);
    dom.sheetView.classList.toggle('hidden', !isSheet);
    [...dom.tabs.children].forEach((tab) => tab.classList.toggle('active', tab.dataset.view === view));
    hideChartTooltip();
    syncHash(view);
    if (isSheet) {
      if (!state.selectedCell || state.selectedCell.sheetName !== view) {
        state.selectedCell = { sheetName: view, row: 0, col: 0 };
      }
      if (state.sheetSelection.sheetName !== view) state.sheetSelection = { sheetName: view, rows: new Set(), anchor: null };
      renderSheetGrid();
      updateSelectionUI();
      renderSelectionBar(dom.sheetActions, sheetSelectionKeys(), 'sheet');
    } else {
      refreshViews();
      if (view === VIEW.consolidated) renderSelectionBar(dom.consolidatedActions, [...state.selectedKeys], 'consolidated');
    }
  }

  function renderSheetGrid() {
    const sheetName = state.activeView;
    if (SPECIAL_VIEWS.has(sheetName)) return;
    const matrix = state.matrices[sheetName] || [[]];
    let dimensions = { height: matrix.length, width: Math.max(1, ...matrix.map((row) => row.length)) };
    try { dimensions = state.hf.getSheetDimensions(state.sheetIds[sheetName]); } catch { /* Use matrix dimensions. */ }
    const rows = Math.min(Math.max(dimensions.height, 20), MAX_GRID_ROWS);
    const cols = Math.min(Math.max(dimensions.width, 12), MAX_GRID_COLS);
    const selectedRows = state.sheetSelection.sheetName === sheetName ? state.sheetSelection.rows : new Set();

    const table = document.createElement('table');
    table.className = 'spreadsheet-table';
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headerRow.append(document.createElement('th'));
    for (let col = 0; col < cols; col += 1) {
      const th = document.createElement('th');
      th.textContent = columnName(col);
      headerRow.append(th);
    }
    thead.append(headerRow);
    table.append(thead);

    const tbody = document.createElement('tbody');
    for (let row = 0; row < rows; row += 1) {
      const tr = document.createElement('tr');
      if (selectedRows.has(row)) tr.classList.add('row-selected');
      const numberCell = document.createElement('td');
      numberCell.className = 'row-number';
      numberCell.textContent = String(row + 1);
      numberCell.tabIndex = 0;
      numberCell.setAttribute('role', 'button');
      numberCell.setAttribute('aria-pressed', String(selectedRows.has(row)));
      numberCell.setAttribute('aria-label', `Select row ${row + 1}`);
      numberCell.addEventListener('click', (event) => toggleSheetRow(sheetName, row, event));
      numberCell.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          toggleSheetRow(sheetName, row, { ctrlKey: true, metaKey: false, shiftKey: event.shiftKey });
        }
      });
      tr.append(numberCell);
      for (let col = 0; col < cols; col += 1) {
        const td = document.createElement('td');
        const raw = matrix[row]?.[col] ?? null;
        const value = readCalculatedValue(sheetName, row, col);
        td.textContent = displayValue(value);
        td.translate = false;
        td.title = raw === null ? '' : String(raw);
        td.dataset.row = String(row);
        td.dataset.col = String(col);
        if (typeof raw === 'string' && raw.startsWith('=')) td.classList.add('formula');
        if (isFormulaError(value)) td.classList.add('error');
        if (state.selectedCell?.sheetName === sheetName && state.selectedCell.row === row && state.selectedCell.col === col) td.classList.add('selected');
        td.addEventListener('click', () => selectCell(sheetName, row, col));
        td.addEventListener('dblclick', () => {
          selectCell(sheetName, row, col);
          dom.formulaInput.focus();
          dom.formulaInput.select();
        });
        tr.append(td);
      }
      tbody.append(tr);
    }
    table.append(tbody);
    dom.grid.replaceChildren(table);
  }

  function toggleSheetRow(sheetName, row, event) {
    const selection = state.sheetSelection.sheetName === sheetName
      ? state.sheetSelection
      : (state.sheetSelection = { sheetName, rows: new Set(), anchor: null });
    if (event.shiftKey && selection.anchor !== null) {
      const [from, to] = [selection.anchor, row].sort((a, b) => a - b);
      if (!(event.ctrlKey || event.metaKey)) selection.rows.clear();
      for (let index = from; index <= to; index += 1) selection.rows.add(index);
    } else if (event.ctrlKey || event.metaKey) {
      if (selection.rows.has(row)) selection.rows.delete(row);
      else selection.rows.add(row);
      selection.anchor = row;
    } else {
      const only = selection.rows.size === 1 && selection.rows.has(row);
      selection.rows.clear();
      if (!only) selection.rows.add(row);
      selection.anchor = row;
    }
    dom.grid.querySelectorAll('tbody tr').forEach((tr, index) => {
      tr.classList.toggle('row-selected', selection.rows.has(index));
      tr.firstElementChild?.setAttribute('aria-pressed', String(selection.rows.has(index)));
    });
    renderSelectionBar(dom.sheetActions, sheetSelectionKeys(), 'sheet');
  }

  // Part keys for the selected rows of the active kit tab.
  function sheetSelectionKeys() {
    const { sheetName, rows } = state.sheetSelection;
    const mapping = state.mappings[sheetName];
    if (!sheetName || !mapping || mapping.part === null || !rows.size) return [];
    const keys = [];
    [...rows].sort((a, b) => a - b).forEach((row) => {
      if (row < mapping.dataStartRow) return;
      const part = readCalculatedValue(sheetName, row, mapping.part);
      const key = part === null || part === undefined ? '' : normalizePart(String(part).trim());
      if (key && !/^(?:total(?:s)?|effective\b|confidential\b)/i.test(key)) keys.push(key);
    });
    return [...new Set(keys)];
  }

  function selectCell(sheetName, row, col) {
    state.selectedCell = { sheetName, row, col };
    dom.grid.querySelectorAll('td.selected').forEach((cell) => cell.classList.remove('selected'));
    const cell = dom.grid.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
    if (cell) cell.classList.add('selected');
    updateSelectionUI();
  }

  function updateSelectionUI() {
    if (!state.selectedCell) return;
    const { sheetName, row, col } = state.selectedCell;
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    const raw = state.matrices[sheetName]?.[row]?.[col] ?? '';
    dom.selectedAddress.textContent = address;
    dom.formulaInput.value = raw === null ? '' : String(raw).replace(/^'(?==)/, '');
    dom.targetRange.value = `${address}:${address}`;
    if (typeof raw === 'string' && raw.startsWith('=')) dom.rangeFormula.value = raw;
    renderDependencies(sheetName, row, col);
  }

  function saveSelectedCell() {
    if (!state.selectedCell) return;
    const raw = parseUserInput(dom.formulaInput.value);
    const { sheetName, row, col } = state.selectedCell;
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    const previous = captureBlock(sheetName, row, col, [[undefined]]);
    try {
      pushHistory(`Cell ${sheetName}!${address} updated`, () => {
        setEngineBlock(sheetName, row, col, [[raw]]);
        afterWorkbookEdit(sheetName);
      }, { restore: () => restoreBlock(sheetName, row, col, previous) });
      showToast(`Updated ${sheetName}!${address}.`);
    } catch (error) {
      showToast(error.message || 'The cell could not be updated.', true);
    }
  }

  function applyToRange(clear) {
    if (!state.activeView || SPECIAL_VIEWS.has(state.activeView)) return;
    try {
      const range = parseRange(dom.targetRange.value);
      const cellCount = (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
      if (cellCount > MAX_FILL_CELLS) throw new Error(`The range is too large. The limit is ${formatNumber(MAX_FILL_CELLS, 0)} cells.`);
      const seed = clear ? null : parseUserInput(dom.rangeFormula.value);
      const block = [];
      for (let row = range.s.r; row <= range.e.r; row += 1) {
        const values = [];
        for (let col = range.s.c; col <= range.e.c; col += 1) {
          values.push(typeof seed === 'string' && seed.startsWith('=')
            ? shiftFormula(seed, row - range.s.r, col - range.s.c)
            : seed);
        }
        block.push(values);
      }
      const previous = captureBlock(state.activeView, range.s.r, range.s.c, block);
      pushHistory(clear ? `${state.activeView}!${dom.targetRange.value} cleared` : `${state.activeView}!${dom.targetRange.value} filled`, () => {
        setEngineBlock(state.activeView, range.s.r, range.s.c, block);
        state.selectedCell = { sheetName: state.activeView, row: range.s.r, col: range.s.c };
        afterWorkbookEdit(state.activeView);
      }, { restore: () => restoreBlock(state.activeView, range.s.r, range.s.c, previous) });
      showToast(clear ? 'Range cleared. Ctrl+Z restores it.' : `Formula applied to ${formatNumber(cellCount, 0)} cells. Ctrl+Z restores them.`);
    } catch (error) {
      showToast(error.message || 'The range could not be updated.', true);
    }
  }

  function parseRange(value) {
    const cleaned = String(value).trim().replace(/\$/g, '').toUpperCase();
    if (!/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/.test(cleaned)) throw new Error('Use an A1 range such as G2:G20 or B2:F12.');
    return XLSX.utils.decode_range(cleaned);
  }

  function shiftFormula(formula, rowOffset, colOffset) {
    const referencePattern = /((?:'[^']*(?:''[^']*)*'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)/g;
    return formula.replace(referencePattern, (match, sheet, fixedCol, letters, fixedRow, rowText) => {
      const originalCol = XLSX.utils.decode_col(letters);
      const originalRow = Number(rowText) - 1;
      const nextCol = fixedCol ? originalCol : originalCol + colOffset;
      const nextRow = fixedRow ? originalRow : originalRow + rowOffset;
      if (nextCol < 0 || nextRow < 0) return '#REF!';
      return `${sheet || ''}${fixedCol || ''}${columnName(nextCol)}${fixedRow || ''}${nextRow + 1}`;
    });
  }

  function setEngineBlock(sheetName, startRow, startCol, block) {
    const sheetId = state.sheetIds[sheetName];
    state.hf.setCellContents({ sheet: sheetId, row: startRow, col: startCol }, block);
    const matrix = state.matrices[sheetName];
    block.forEach((values, rowOffset) => {
      const row = startRow + rowOffset;
      if (!matrix[row]) matrix[row] = [];
      values.forEach((value, colOffset) => {
        const col = startCol + colOffset;
        matrix[row][col] = value;
        updateWorksheetCell(sheetName, row, col, value);
      });
    });
  }

  /* ---------- Worksheet undo helpers ---------- */

  // Reads the current raw values of a block so an edit can be reversed cell by cell.
  function captureBlock(sheetName, startRow, startCol, block) {
    const matrix = state.matrices[sheetName] || [];
    return block.map((values, rowOffset) => values.map((_, colOffset) => {
      const row = startRow + rowOffset;
      const col = startCol + colOffset;
      const value = matrix[row]?.[col];
      return value === undefined ? null : value;
    }));
  }

  // Puts previously captured values back, rendering once at the end rather than per cell.
  function restoreBlock(sheetName, startRow, startCol, values) {
    const block = values.map((row) => [...row]);
    setEngineBlock(sheetName, startRow, startCol, block);
    afterWorkbookEdit(sheetName);
  }

  function updateWorksheetCell(sheetName, row, col, value) {
    const worksheet = state.workbook.Sheets[sheetName];
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    const previous = worksheet[address] || {};
    const preserved = {};
    ['s', 'z', 'l', 'c'].forEach((key) => { if (previous[key] !== undefined) preserved[key] = previous[key]; });

    if (value === null || value === '') {
      delete worksheet[address];
    } else if (typeof value === 'string' && value.startsWith('=')) {
      worksheet[address] = { ...preserved, t: 'n', f: value.slice(1) };
    } else if (value instanceof Date) {
      worksheet[address] = { ...preserved, t: 'd', v: value };
    } else if (typeof value === 'number') {
      worksheet[address] = { ...preserved, t: 'n', v: value };
    } else if (typeof value === 'boolean') {
      worksheet[address] = { ...preserved, t: 'b', v: value };
    } else {
      worksheet[address] = { ...preserved, t: 's', v: String(value).replace(/^'(?==)/, '') };
    }

    const current = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: row, c: col }, e: { r: row, c: col } };
    current.s.r = Math.min(current.s.r, row);
    current.s.c = Math.min(current.s.c, col);
    current.e.r = Math.max(current.e.r, row);
    current.e.c = Math.max(current.e.c, col);
    worksheet['!ref'] = XLSX.utils.encode_range(current);
  }

  function afterWorkbookEdit(sheetName) {
    markProjectDirty();
    rebuildDependencyIndex();
    rebuildConsolidation();
    if (state.activeView === sheetName) {
      renderSheetGrid();
      updateSelectionUI();
    }
  }

  function importNamedExpressions(workbook) {
    const names = workbook.Workbook?.Names || [];
    names.forEach((entry) => {
      if (!entry.Name || !entry.Ref || entry.Name.startsWith('_xlnm.')) return;
      const expression = String(entry.Ref).startsWith('=') ? String(entry.Ref) : `=${entry.Ref}`;
      try {
        state.hf.addNamedExpression(entry.Name, expression);
        state.variables.push({ name: entry.Name, expression });
      } catch (error) {
        console.warn(`Named expression ${entry.Name} was not imported:`, error);
      }
    });
  }

  function addVariable(event) {
    event.preventDefault();
    const name = dom.variableName.value.trim();
    const expressionText = dom.variableExpression.value.trim();
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) {
      showToast('Variable names must start with a letter or underscore and contain no spaces.', true);
      return;
    }
    if (!expressionText) {
      showToast('Enter a value or formula for the variable.', true);
      return;
    }
    try {
      const expression = parseUserInput(expressionText);
      pushHistory(`Variable ${name} added`, () => {
        state.hf.addNamedExpression(name, expression);
        state.variables.push({ name, expression: expressionText });
        markProjectDirty();
        renderVariables();
        renderSheetGrid();
        rebuildConsolidation();
      });
      dom.variableForm.reset();
    } catch (error) {
      showToast(error.message || 'The variable could not be added.', true);
    }
  }

  function renderVariables() {
    dom.variablesList.replaceChildren();
    if (!state.variables.length) {
      const note = document.createElement('p');
      note.className = 'fine-print';
      note.textContent = 'No workbook variables are defined.';
      dom.variablesList.append(note);
      return;
    }
    state.variables.forEach((variable) => {
      const row = document.createElement('div');
      row.className = 'variable-row';
      const label = document.createElement('span');
      label.textContent = `${variable.name} = ${variable.expression}`;
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'icon-button';
      remove.textContent = '×';
      remove.title = `Remove ${variable.name}`;
      remove.addEventListener('click', () => removeVariable(variable.name));
      row.append(label, remove);
      dom.variablesList.append(row);
    });
  }

  function removeVariable(name) {
    pushHistory(`Variable ${name} removed`, () => {
      state.hf.removeNamedExpression(name);
      state.variables = state.variables.filter((variable) => variable.name !== name);
      markProjectDirty();
      renderVariables();
      renderSheetGrid();
      rebuildConsolidation();
    });
  }

  function rebuildDependencyIndex() {
    const index = new Map();
    Object.entries(state.matrices).forEach(([sheetName, matrix]) => {
      matrix.forEach((row, rowIndex) => {
        (row || []).forEach((raw, colIndex) => {
          if (typeof raw !== 'string' || !raw.startsWith('=')) return;
          const dependent = `${sheetName}!${XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })}`;
          extractFormulaReferences(raw, sheetName).forEach((reference) => {
            if (!index.has(reference.key)) index.set(reference.key, new Set());
            index.get(reference.key).add(dependent);
          });
        });
      });
    });
    state.dependencyIndex = index;
  }

  function extractFormulaReferences(formula, currentSheet) {
    const references = [];
    const pattern = /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/g;
    let match;
    while ((match = pattern.exec(formula)) !== null) {
      const sheetName = (match[1] ? match[1].replace(/''/g, "'") : match[2]) || currentSheet;
      if (!(sheetName in state.matrices)) continue;
      const start = XLSX.utils.decode_cell(match[3].replace(/\$/g, ''));
      const end = match[4] ? XLSX.utils.decode_cell(match[4].replace(/\$/g, '')) : start;
      const size = (end.r - start.r + 1) * (end.c - start.c + 1);
      if (size > 2000 || end.r < start.r || end.c < start.c) continue;
      for (let row = start.r; row <= end.r; row += 1) {
        for (let col = start.c; col <= end.c; col += 1) {
          const address = XLSX.utils.encode_cell({ r: row, c: col });
          references.push({ sheetName, row, col, address, key: `${sheetName}!${address}` });
        }
      }
    }
    return uniqueBy(references, (reference) => reference.key);
  }

  function renderDependencies(sheetName, row, col) {
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    const raw = state.matrices[sheetName]?.[row]?.[col] ?? null;
    const key = `${sheetName}!${address}`;
    const precedents = typeof raw === 'string' && raw.startsWith('=') ? extractFormulaReferences(raw, sheetName) : [];
    const dependents = [...(state.dependencyIndex.get(key) || [])];

    dom.dependencyPanel.replaceChildren();
    const eyebrow = document.createElement('p');
    eyebrow.className = 'eyebrow';
    eyebrow.textContent = 'Cell links';
    const title = document.createElement('h3');
    title.textContent = key;
    dom.dependencyPanel.append(eyebrow, title);

    if (typeof raw === 'string' && raw.startsWith('=')) {
      const formula = document.createElement('p');
      formula.className = 'dependency-formula';
      formula.textContent = raw;
      dom.dependencyPanel.append(formula);
    }

    dom.dependencyPanel.append(dependencySection('Uses', precedents.map((item) => item.key)));
    dom.dependencyPanel.append(dependencySection('Used by', dependents));
  }

  function dependencySection(label, keys) {
    const section = document.createElement('section');
    const heading = document.createElement('h3');
    heading.textContent = label;
    section.append(heading);
    if (!keys.length) {
      const empty = document.createElement('p');
      empty.className = 'muted';
      empty.textContent = 'None detected.';
      section.append(empty);
      return section;
    }
    const list = document.createElement('div');
    list.className = 'dependency-list';
    keys.slice(0, 100).forEach((key) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'dependency-link';
      button.textContent = key;
      button.addEventListener('click', () => jumpToCell(key));
      list.append(button);
    });
    section.append(list);
    return section;
  }

  function jumpToCell(key) {
    const divider = key.lastIndexOf('!');
    const sheetName = key.slice(0, divider);
    const address = key.slice(divider + 1);
    const point = XLSX.utils.decode_cell(address);
    showView(sheetName);
    state.selectedCell = { sheetName, row: point.r, col: point.c };
    renderSheetGrid();
    updateSelectionUI();
    const cell = dom.grid.querySelector(`td[data-row="${point.r}"][data-col="${point.c}"]`);
    cell?.scrollIntoView({ block: 'center', inline: 'center' });
  }

  async function exportSecureProject() {
    if (!state.workbook) {
      showToast('Import a workbook before exporting a secure project.', true);
      return;
    }
    const password = dom.securePassword.value;
    if (password.length < 10) {
      showToast('Use a password with at least 10 characters.', true);
      return;
    }
    if (password !== dom.securePasswordConfirm.value) {
      showToast('The secure project passwords do not match.', true);
      return;
    }
    if (!window.crypto?.subtle) {
      showToast('Secure project encryption is not available in this browser.', true);
      return;
    }

    dom.secureExport.disabled = true;
    setStatus('Encrypting the secure project…', 'busy');
    try {
      const payload = {
        version: 1,
        createdAt: new Date().toISOString(),
        fileName: state.fileName,
        workbook: XLSX.write(state.workbook, { type: 'base64', bookType: 'xlsx', cellStyles: true, compression: true }),
        config: state.config,
        warehouse: state.warehouse,
        mappings: state.mappings,
        variables: state.variables,
        wireRelations: state.wireRelations,
        settings: collectSettings(),
      };
      const salt = crypto.getRandomValues(new Uint8Array(16));
      const iv = crypto.getRandomValues(new Uint8Array(12));
      const iterations = 250000;
      const key = await deriveSecureKey(password, salt, iterations, ['encrypt']);
      const plaintext = new TextEncoder().encode(JSON.stringify(payload));
      const ciphertext = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, plaintext));
      const envelope = {
        format: 'partslist-secure',
        version: 1,
        kdf: { name: 'PBKDF2', hash: 'SHA-256', iterations, salt: bytesToBase64(salt) },
        cipher: { name: 'AES-GCM', iv: bytesToBase64(iv) },
        data: bytesToBase64(ciphertext),
      };
      const base = state.fileName.replace(/\.[^.]+$/, '') || 'partslist';
      const secureBlob = new Blob([JSON.stringify(envelope)], { type: 'application/json' });
      prepareSecureDownload(secureBlob, `${base}.partslist`);
      dom.securePassword.value = '';
      dom.securePasswordConfirm.value = '';
      setStatus(`Encrypted project is ready to download for ${state.fileName}.`);
      showToast('Encrypted project ready. Use the download button.');
    } catch (error) {
      console.error(error);
      setStatus('The secure project could not be encrypted.', 'error');
      showToast(error.message || 'Secure project export failed.', true);
    } finally {
      dom.secureExport.disabled = false;
    }
  }

  async function importSecureProject(file) {
    const password = dom.securePassword.value;
    if (!password) {
      showToast('Enter the secure project password first.', true);
      return;
    }
    if (!window.crypto?.subtle) {
      showToast('Secure project decryption is not available in this browser.', true);
      return;
    }
    setStatus(`Decrypting ${file.name}…`, 'busy');
    try {
      const envelope = JSON.parse(await file.text());
      if (envelope.format !== 'partslist-secure' || envelope.version !== 1) throw new Error('This is not a supported PartsList secure project.');
      const iterations = Number(envelope.kdf?.iterations);
      if (envelope.kdf?.name !== 'PBKDF2' || envelope.kdf?.hash !== 'SHA-256' || !Number.isInteger(iterations) || iterations < 100000 || iterations > 1000000) throw new Error('The secure project uses unsupported key settings.');
      if (envelope.cipher?.name !== 'AES-GCM') throw new Error('The secure project uses an unsupported cipher.');
      const salt = base64ToBytes(envelope.kdf.salt);
      const iv = base64ToBytes(envelope.cipher?.iv);
      const ciphertext = base64ToBytes(envelope.data);
      if (salt.length !== 16 || iv.length !== 12 || ciphertext.length < 17) throw new Error('The secure project envelope is damaged.');
      const key = await deriveSecureKey(password, salt, iterations, ['decrypt']);
      let plaintext;
      try {
        plaintext = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, ciphertext);
      } catch {
        throw new Error('The password is incorrect or the secure project is damaged.');
      }
      const payload = JSON.parse(new TextDecoder().decode(plaintext));
      if (payload.version !== 1 || typeof payload.workbook !== 'string') throw new Error('The decrypted project data is invalid.');
      const workbookBytes = base64ToBytes(payload.workbook);
      const workbookFile = new File([workbookBytes], payload.fileName || 'secure-project.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      await importWorkbook(workbookFile);

      state.config = { ...state.config, ...(payload.config || {}) };
      state.warehouse = {
        ...WAREHOUSE_DEFAULTS,
        ...(payload.warehouse || {}),
        rates: { ...WAREHOUSE_RATE_DEFAULTS, ...(payload.warehouse?.rates || {}) },
      };
      state.warehouse.unitsPerShelf = state.warehouse.binsPerShelf * state.warehouse.unitsPerBin;
      state.wireRelations = Array.isArray(payload.wireRelations) ? payload.wireRelations.filter(isValidWireRelation) : [];
      if (payload.settings) applySettings(payload.settings);
      state.pendingDisabledKits = null;
      if (payload.mappings && typeof payload.mappings === 'object') {
        state.workbook.SheetNames.forEach((sheetName) => {
          if (payload.mappings[sheetName]) state.mappings[sheetName] = { ...state.mappings[sheetName], ...payload.mappings[sheetName] };
        });
      }
      restoreVariables(payload.variables);
      syncControlsFromState();
      persistWarehouseRates();
      saveSettingsSoon();
      renderMappings();
      renderVariables();
      renderGroups();
      rebuildConsolidation();
      renderTabs();
      showView(VIEW.consolidated);
      dom.securePassword.value = '';
      dom.securePasswordConfirm.value = '';
      setStatus(`Secure project unlocked: ${state.fileName}.`);
      showToast('Secure project imported and unlocked.');
      await refreshRates(false);
      markProjectSaved('Encrypted project opened');
    } catch (error) {
      console.error(error);
      setStatus(`Could not open ${file.name}: ${error.message}`, 'error');
      showToast(error.message || 'Secure project import failed.', true);
    }
  }

  function syncControlsFromState(active = null) {
    dom.sourceCurrency.value = state.config.sourceCurrency;
    dom.outputCurrency.value = state.config.outputCurrency;
    dom.warehouseCountry.value = state.warehouse.country;
    dom.languageSelect.value = state.language;
    writeInputValue(dom.whShelfEnabled, state.warehouse.shelfEnabled);
    writeInputValue(dom.whDrawerEnabled, state.warehouse.drawerEnabled);
    writeInputValue(dom.whPalletEnabled, state.warehouse.palletEnabled);
    if (dom.whPlannedPallets !== active) writeInputValue(dom.whPlannedPallets, state.warehouse.plannedPallets);
    const bindings = [
      [dom.whShelfShare, 'shelfShare'], [dom.whDrawerShare, 'drawerShare'],
      [dom.whBinsPerShelf, 'binsPerShelf'], [dom.whShelvesPerRack, 'shelvesPerRack'],
      [dom.whUnitsPerBin, 'unitsPerBin'], [dom.whUnitsShelf, 'unitsPerShelf'],
      [dom.whUnitsDrawer, 'unitsPerDrawer'], [dom.whUnitsPallet, 'unitsPerPallet'],
      [dom.whReceipts, 'receipts'], [dom.whReceiptLines, 'receiptLines'],
      [dom.whOrders, 'orders'], [dom.whOrderLines, 'orderLines'],
      [dom.whEdiLabels, 'ediLabels'], [dom.whBusinessParcels, 'businessParcels'],
      [dom.whPrivateParcels, 'privateParcels'], [dom.whPackaging, 'packaging'],
    ];
    bindings.forEach(([input, key]) => { if (input !== active) input.value = String(state.warehouse[key]); });
    dom.whPalletType.value = state.warehouse.palletType;
    dom.whPalletHeight.value = state.warehouse.palletHeight;
    dom.whWms.checked = Boolean(state.warehouse.includeWms);
    dom.warehouseRateInputs.querySelectorAll('[data-warehouse-rate]').forEach((input) => {
      if (input !== active) input.value = String(state.warehouse.rates[input.dataset.warehouseRate] ?? 0);
    });
  }

  function restoreVariables(variables) {
    if (!Array.isArray(variables)) return;
    const existing = new Set(state.variables.map((variable) => variable.name.toUpperCase()));
    variables.forEach((variable) => {
      if (!variable || !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(variable.name) || existing.has(variable.name.toUpperCase())) return;
      try {
        state.hf.addNamedExpression(variable.name, parseUserInput(variable.expression));
        state.variables.push({ name: variable.name, expression: String(variable.expression) });
        existing.add(variable.name.toUpperCase());
      } catch { /* Invalid or unsupported stored named expressions are ignored. */ }
    });
  }

  function isValidWireRelation(edge) {
    return edge && typeof edge.id === 'string' && typeof edge.from === 'string' && typeof edge.to === 'string' && typeof edge.label === 'string';
  }

  async function deriveSecureKey(password, salt, iterations, usages) {
    const material = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations, hash: 'SHA-256' }, material, { name: 'AES-GCM', length: 256 }, false, usages);
  }

  function bytesToBase64(bytes) {
    let binary = '';
    const chunkSize = 0x8000;
    for (let offset = 0; offset < bytes.length; offset += chunkSize) binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
    return btoa(binary);
  }

  function base64ToBytes(value) {
    const binary = atob(String(value || ''));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  function prepareSecureDownload(blob, fileName) {
    if (state.secureDownloadUrl) URL.revokeObjectURL(state.secureDownloadUrl);
    state.secureDownloadUrl = URL.createObjectURL(blob);
    dom.secureDownload.href = state.secureDownloadUrl;
    dom.secureDownload.download = fileName;
    dom.secureDownload.classList.remove('hidden');
  }

  function exportWorkbook() {
    if (!state.workbook || !state.consolidated.length) {
      showToast('There are no consolidated parts to export.', true);
      return;
    }
    try {
      // Sheets generated by an earlier export are rebuilt, not copied, so names stay stable across round trips.
      const sourceNames = state.workbook.SheetNames.filter((name) => !isGeneratedSheet(state.workbook.Sheets[name]));
      const usedNames = new Set(sourceNames);
      const claim = (base) => {
        const name = uniqueSheetName(base, usedNames);
        usedNames.add(name);
        return name;
      };
      const consolidatedName = claim('Consolidated');
      const profitabilityName = claim('Profitability');
      const warehousingName = claim('Warehousing');
      const assumptionsName = claim('Assumptions');
      const projectDataName = claim('Project data');
      const settingsName = claim('PartsList settings');
      const output = XLSX.utils.book_new();
      const consolidated = buildConsolidatedSheet(assumptionsName);
      XLSX.utils.book_append_sheet(output, consolidated.sheet, consolidatedName);
      XLSX.utils.book_append_sheet(output, buildProfitabilitySheet(consolidatedName, assumptionsName, consolidated.layout), profitabilityName);
      XLSX.utils.book_append_sheet(output, buildWarehousingSheet(), warehousingName);
      XLSX.utils.book_append_sheet(output, buildAssumptionsSheet(), assumptionsName);
      XLSX.utils.book_append_sheet(output, buildProjectDataSheet(), projectDataName);
      sourceNames.forEach((sheetName) => {
        XLSX.utils.book_append_sheet(output, state.workbook.Sheets[sheetName], sheetName);
      });
      XLSX.utils.book_append_sheet(output, buildSettingsSheet(), settingsName);
      output.Workbook = output.Workbook || {};
      output.Workbook.Sheets = output.SheetNames.map((name) => ({ name, Hidden: name === settingsName ? 1 : 0 }));
      output.Workbook.CalcPr = { calcMode: 'auto', fullCalcOnLoad: '1', forceFullCalc: '1' };
      if (state.workbook.Workbook?.Names) output.Workbook.Names = state.workbook.Workbook.Names;
      if (state.workbook.Props) output.Props = { ...state.workbook.Props };
      if (state.workbook.Custprops) output.Custprops = { ...state.workbook.Custprops };
      const base = state.fileName.replace(/\.[^.]+$/, '').replace(/-(?:consolidated|partslist-project)$/i, '') || 'partslist';
      XLSX.writeFile(output, `${base}-partslist-project.xlsx`, { bookType: 'xlsx', cellStyles: true, compression: true });
      markProjectSaved('Project workbook exported');
      showToast('Project workbook saved. Re-import it to restore source sheets, variables, mappings, warehouse rates and assumptions.');
    } catch (error) {
      console.error(error);
      showToast(error.message || 'The workbook could not be exported.', true);
    }
  }

  function buildAssumptionsSheet() {
    const output = state.config.outputCurrency;
    // Rows 2-8 keep their original addresses so older exports still re-import.
    const rows = [
      ['Assumption', 'Value'],
      ['Discount', state.config.discount],
      ['Price multiplier', state.config.multiplier],
      ['Freight margin', state.config.shippingMargin],
      ['Output currency', output],
      ['USD to output currency', output === 'USD' ? 1 : state.rates.USD ?? null],
      ['NOK to output currency', output === 'NOK' ? 1 : state.rates.NOK ?? null],
      ['Exported at', new Date().toISOString()],
      ['VAT rate', salesVatRate()],
      ['Customs duty', state.freight.dutyRate],
      ['Insurance', state.freight.insuranceRate],
      [`Consolidated shipment (${output})`, toOutput(state.freight.consolidatedShipment, state.freight.consolidatedCurrency) ?? 0],
      [`Clearance & broker fees (${output})`, toOutput(state.freight.clearanceFee, state.freight.clearanceCurrency) ?? 0],
      ['Shipment split basis', { value: 'Purchase value', quantity: 'Quantity', lines: 'Item lines' }[state.freight.allocation]],
      ['Consolidated shipment as entered', `${state.freight.consolidatedShipment} ${state.freight.consolidatedCurrency}`],
      ['Clearance fees as entered', `${state.freight.clearanceFee} ${state.freight.clearanceCurrency}`],
      [],
      ['Rate source', 'https://frankfurter.dev/'],
      ['Rate type', 'Daily reference / mid-market rate'],
      [],
      ['Source currency', `Rate to ${output}`, 'Rate date'],
    ];
    const headerRow = rows.length - 1;
    const currencies = new Set([
      ...state.consolidated.map((item) => item.currency),
      state.freight.consolidatedCurrency,
      state.freight.clearanceCurrency,
      ...state.workbook.SheetNames.filter((name) => state.mappings[name]?.enabled).flatMap((name) => [kitSettings(name).shippingCurrency, kitCurrency(kitSettings(name))]),
    ]);
    [...currencies].filter(Boolean).sort().forEach((currency) => {
      rows.push([currency, currency === output ? 1 : state.rates[currency] ?? null, state.rateDates[currency] ?? null]);
    });
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = [{ wch: 34 }, { wch: 30 }, { wch: 14 }];
    styleHeader(sheet, 0, 0, 1);
    styleHeader(sheet, headerRow, 0, 2);
    ['B2', 'B4', 'B9', 'B10', 'B11'].forEach((address) => { if (sheet[address]) sheet[address].z = '0.0%'; });
    ['B12', 'B13'].forEach((address) => { if (sheet[address]) sheet[address].z = '#,##0.00'; });
    return sheet;
  }

  const CONSOLIDATED_EXPORT_COLUMNS = ['sources', 'part', 'description', 'altPart', 'quantity', 'amountOut', 'match', 'group', 'salesPerYear', 'currency', 'unitPrice', 'fx', 'unitOut', 'discount', 'purchase', 'multiplier', 'sales', 'kitFreight', 'consFreight', 'importCosts', 'freightImport', 'margin', 'freightWithMargin', 'lineTotal', 'vat', 'lineTotalVat', 'landed', 'pricingSource'];

  function buildConsolidatedSheet(assumptionsName) {
    const o = state.config.outputCurrency;
    const A = quoteSheetName(assumptionsName);
    const L = Object.fromEntries(CONSOLIDATED_EXPORT_COLUMNS.map((key, index) => [key, XLSX.utils.encode_col(index)]));
    const lastCol = CONSOLIDATED_EXPORT_COLUMNS.length - 1;
    const headerLabels = {
      sources: 'Source Sheets', part: 'Part', description: 'Description', altPart: 'Alt. Part No.', quantity: 'Quantity', amountOut: `Amount in ${o}`,
      match: 'Match', group: 'Sales Group', salesPerYear: 'Expected Sales / yr', currency: 'Currency', unitPrice: 'Unit Price', fx: `FX to ${o}`,
      unitOut: `Unit in ${o}`, discount: 'Discount', purchase: `Purchase After Discount in ${o}`, multiplier: 'Price Multiplier', sales: `Sales Price in ${o}`,
      kitFreight: `Kit Freight in ${o}`, consFreight: `Consolidated Shipment in ${o}`, importCosts: `Duty, Insurance & Fees in ${o}`, freightImport: `Freight & Import in ${o}`,
      margin: 'Freight Margin', freightWithMargin: `Freight & Import incl. Margin in ${o}`, lineTotal: `Line Total excl. VAT in ${o}`, vat: `VAT in ${o}`,
      lineTotalVat: `Line Total incl. VAT in ${o}`, landed: `Landed Cost in ${o}`, pricingSource: 'Pricing Source',
    };
    const rows = [
      ['Consolidated parts'],
      ['Common parts are highlighted in yellow and listed first. Quantity is the maximum found (or the stock override); pricing inputs come from the row with that maximum quantity. Excluded parts are not listed.'],
      CONSOLIDATED_EXPORT_COLUMNS.map((key) => headerLabels[key]),
    ];
    const groups = [
      ['Common', state.consolidated.filter((item) => item.common)],
      ['Unique', state.consolidated.filter((item) => !item.common)],
    ];
    const layout = [];
    let excelRow = rows.length + 1;
    groups.forEach(([label, items]) => {
      if (!items.length) return;
      layout.push({ label, row: excelRow });
      excelRow += 1;
      items.forEach((item) => {
        layout.push({ item, row: excelRow });
        excelRow += 1;
      });
    });
    const itemRows = layout.filter((entry) => entry.item).map((entry) => entry.row);
    const first = Math.min(...itemRows);
    const last = Math.max(...itemRows);
    const range = (key) => `${L[key]}$${first}:${L[key]}$${last}`;
    const share = (r) => {
      if (state.freight.allocation === 'lines') return `IF(${L.quantity}${r}>0,1,0)/MAX(1,COUNTIF(${range('quantity')},">0"))`;
      const key = state.freight.allocation === 'quantity' ? 'quantity' : 'purchase';
      return `IF(SUM(${range(key)})=0,0,${L[key]}${r}/SUM(${range(key)}))`;
    };
    // The shared shipment, insurance and clearance cells describe one pool of imports at the project-wide duty.
    const uniformPools = state.consolidated.every((item) => !makerHasOwnShipment(item.manufacturer) && makerInsuranceRate(item.manufacturer) === state.freight.insuranceRate);
    const uniformImports = uniformPools && state.consolidated.every((item) => item.routeTreatment === 'import' && Math.abs(item.dutyRate - state.freight.dutyRate) < 1e-12);
    const commonRows = [];
    const groupRows = [];

    layout.forEach((entry) => {
      if (!entry.item) {
        rows.push([`${entry.label} parts`]);
        groupRows.push(entry.row - 1);
        return;
      }
      const { item, row: r } = entry;
      const cells = {
        sources: item.sources.join(', '),
        part: item.part,
        description: item.description,
        altPart: item.altParts.join(', '),
        quantity: item.maxQuantity,
        amountOut: formulaCell(`${L.quantity}${r}*${L.unitOut}${r}`, item.convertedTotal),
        match: item.common ? 'Common' : 'Unique',
        group: item.groupName || '',
        salesPerYear: item.salesPerYear,
        currency: item.currency,
        unitPrice: item.unitPrice,
        fx: item.fx,
        unitOut: formulaCell(`${L.unitPrice}${r}*${L.fx}${r}`, item.convertedUnit),
        // A kit with its own discount (for example a net price list) cannot follow the project-wide cell.
        discount: Math.abs(item.discount - state.config.discount) < 1e-12 ? formulaCell(`${A}!$B$2`, item.discount) : item.discount,
        purchase: formulaCell(`${L.amountOut}${r}*(1-${L.discount}${r})`, item.discountedTotal),
        multiplier: item.multiplierOverridden ? item.multiplier : formulaCell(`${A}!$B$3`, item.multiplier),
        sales: formulaCell(`${L.purchase}${r}*${L.multiplier}${r}`, item.sellingTotal),
        kitFreight: item.kitFreight,
        consFreight: uniformPools ? formulaCell(`${A}!$B$12*${share(r)}`, item.consolidatedFreight) : item.consolidatedFreight,
        // The shared duty, insurance and clearance cells describe imports at the project-wide duty. Lines with their own
        // duty or a route without customs (domestic or intra-EU) are written as values instead of a wrong formula.
        importCosts: uniformImports
          ? formulaCell(`${L.purchase}${r}*${A}!$B$11+(${L.purchase}${r}+${L.kitFreight}${r}+${L.consFreight}${r}+${L.purchase}${r}*${A}!$B$11)*${A}!$B$10+${A}!$B$13*${share(r)}`, item.importCosts)
          : item.importCosts,
        freightImport: formulaCell(`${L.kitFreight}${r}+${L.consFreight}${r}+${L.importCosts}${r}`, item.freightAndImport),
        margin: formulaCell(`${A}!$B$4`, item.shippingMargin),
        freightWithMargin: formulaCell(`${L.freightImport}${r}/(1-${L.margin}${r})`, item.freightWithMargin),
        lineTotal: formulaCell(`IF(${L.unitPrice}${r}="","",${L.sales}${r}+${L.freightWithMargin}${r})`, item.lineTotal),
        vat: formulaCell(`IF(${L.lineTotal}${r}="","",${L.lineTotal}${r}*${A}!$B$9)`, item.vat),
        lineTotalVat: formulaCell(`IF(${L.lineTotal}${r}="","",${L.lineTotal}${r}+${L.vat}${r})`, item.lineTotalInclVat),
        landed: formulaCell(`${L.purchase}${r}+${L.freightImport}${r}`, item.landedCost),
        pricingSource: `${item.pricingSource} row ${item.pricingRow}`,
      };
      rows.push(CONSOLIDATED_EXPORT_COLUMNS.map((key) => cells[key] ?? null));
      if (item.common) commonRows.push(r - 1);
    });

    const sumKeys = ['quantity', 'amountOut', 'salesPerYear', 'purchase', 'sales', 'kitFreight', 'consFreight', 'importCosts', 'freightImport', 'freightWithMargin', 'lineTotal', 'vat', 'lineTotalVat', 'landed'];
    const totalRow = last + 2;
    rows.push([]);
    rows.push(CONSOLIDATED_EXPORT_COLUMNS.map((key, index) => {
      if (index === 0) return 'Total';
      if (!sumKeys.includes(key)) return null;
      const field = { quantity: 'maxQuantity', amountOut: 'convertedTotal', salesPerYear: 'salesPerYear', purchase: 'discountedTotal', sales: 'sellingTotal', kitFreight: 'kitFreight', consFreight: 'consolidatedFreight', importCosts: 'importCosts', freightImport: 'freightAndImport', freightWithMargin: 'freightWithMargin', lineTotal: 'lineTotal', vat: 'vat', lineTotalVat: 'lineTotalInclVat', landed: 'landedCost' }[key];
      const value = state.consolidated.reduce((sum, item) => sum + (item[field] ?? 0), 0);
      return formulaCell(`SUM(${range(key)})`, value);
    }));

    const sheet = XLSX.utils.aoa_to_sheet(rows);
    const widths = { sources: 28, part: 18, description: 28, altPart: 18, quantity: 10, amountOut: 15, match: 10, group: 16, salesPerYear: 12, currency: 9, unitPrice: 12, fx: 12, unitOut: 13, discount: 10, purchase: 18, multiplier: 11, sales: 16, kitFreight: 14, consFreight: 18, importCosts: 18, freightImport: 16, margin: 11, freightWithMargin: 20, lineTotal: 18, vat: 13, lineTotalVat: 18, landed: 16, pricingSource: 22 };
    sheet['!cols'] = CONSOLIDATED_EXPORT_COLUMNS.map((key) => ({ wch: widths[key] || 14 }));
    sheet['!autofilter'] = { ref: `A3:${XLSX.utils.encode_col(lastCol)}${last}` };
    sheet['!freeze'] = { xSplit: 0, ySplit: 3, topLeftCell: 'A4', activePane: 'bottomLeft', state: 'frozen' };
    styleTitle(sheet, 0, lastCol);
    styleHeader(sheet, 2, 0, lastCol);
    commonRows.forEach((rowIndex) => styleCommonRow(sheet, rowIndex, 0, lastCol));
    groupRows.forEach((rowIndex) => styleGroupRow(sheet, rowIndex, 0, lastCol));
    styleGroupRow(sheet, totalRow - 1, 0, lastCol);
    const moneyKeys = ['amountOut', 'unitPrice', 'unitOut', 'purchase', 'sales', 'kitFreight', 'consFreight', 'importCosts', 'freightImport', 'freightWithMargin', 'lineTotal', 'vat', 'lineTotalVat', 'landed'];
    for (let row = 4; row <= totalRow; row += 1) {
      moneyKeys.forEach((key) => { if (sheet[`${L[key]}${row}`]) sheet[`${L[key]}${row}`].z = '#,##0.00'; });
      ['discount', 'margin'].forEach((key) => { if (sheet[`${L[key]}${row}`]) sheet[`${L[key]}${row}`].z = '0.0%'; });
      ['quantity', 'salesPerYear'].forEach((key) => { if (sheet[`${L[key]}${row}`]) sheet[`${L[key]}${row}`].z = '#,##0.##'; });
      if (sheet[`${L.fx}${row}`]) sheet[`${L.fx}${row}`].z = '0.000000';
    }
    return { sheet, layout: { columns: L, rows: layout.filter((entry) => entry.item), first, last } };
  }

  function buildProfitabilitySheet(consolidatedName, assumptionsName, layout) {
    const model = calculateProfitability();
    const C = quoteSheetName(consolidatedName);
    const A = quoteSheetName(assumptionsName);
    const L = layout.columns;
    const o = state.config.outputCurrency;
    const itemStart = 26;
    const itemEnd = itemStart + layout.rows.length - 1;
    const col = (letter) => `${letter}$${itemStart}:${letter}$${itemEnd}`;
    const rows = [
      ['Profitability estimate'],
      [`Expected annual sales, landed cost and logistics in ${o}, excluding VAT. Change the blue inputs in Consolidated or the per-unit values below and Excel recalculates.${model.complete ? '' : ` ESTIMATE INCOMPLETE: missing ${model.missingReasons.join(' and ')}; totals show only calculable values.`}`],
      [],
      ['Per-unit logistics', 'Value', 'Source'],
      [`Storage per stock unit / month (${o})`, model.perUnit.storagePerUnitMonth ?? 0, model.manual.storagePerUnitMonth ? 'Manual' : 'Warehouse monthly storage ÷ stock units'],
      [`Outtake per unit sold (${o})`, model.perUnit.outtakePerUnit ?? 0, model.manual.outtakePerUnit ? 'Manual' : 'Order handling + EDI + packaging ÷ units sold'],
      [`Outbound freight per unit sold (${o})`, model.perUnit.outboundPerUnit ?? 0, model.manual.outboundPerUnit ? 'Manual' : 'Parcel freight ÷ units sold'],
      ['Customer pays outbound freight', state.sales.customerPaysOutbound, ''],
      [`Receiving & WMS / year (${o})`, model.fixed, 'Fixed warehouse costs'],
      ['VAT rate', formulaCell(`${A}!$B$9`, salesVatRate()), ''],
      [],
      ['Summary', 'Value'],
      [`Annual revenue excl. VAT (${o})`, formulaCell(`SUM(${col('J')})`, model.revenue)],
      [`Annual cost of goods, landed (${o})`, formulaCell(`SUM(${col('K')})`, model.cogs)],
      [`Annual gross profit (${o})`, formulaCell('B13-B14', model.gross)],
      [`Storage / year (${o})`, formulaCell(`SUM(${col('M')})`, model.storage)],
      [`Outtake & outbound freight / year (${o})`, formulaCell(`SUM(${col('N')})`, model.outtake + model.outbound)],
      [`Receiving & WMS / year (${o})`, formulaCell('B9', model.fixed)],
      [`Annual net profit (${o})`, formulaCell('B15-B16-B17-B18', model.net)],
      [`Stock investment, landed (${o})`, formulaCell(`SUM(${C}!${L.landed}$${layout.first}:${L.landed}$${layout.last})`, model.investment)],
      ['Payback (months)', formulaCell('IF(OR(B19<=0,B20=0),"",B20/(B19/12))', model.paybackMonths)],
      ['Annual return on stock', formulaCell('IF(B20=0,"",B19/B20)', model.roi)],
      [`Output VAT / year (${o})`, formulaCell('B13*B10', model.outputVat)],
      [],
      ['Part', 'Description', 'Sales Group', 'Stock Qty', 'Expected Sales / yr', `Unit Price excl. VAT (${o})`, 'Unit Landed Cost', 'Unit Margin', 'Margin %', 'Revenue / yr', 'Cost of Goods / yr', 'Gross Profit / yr', 'Storage / yr', 'Outtake & Freight / yr', 'Net Profit / yr'],
    ];
    const rowsByKey = new Map(model.rows.map((row) => [row.item.key, row]));
    layout.rows.forEach(({ item, row: cr }, index) => {
      const r = itemStart + index;
      const profit = rowsByKey.get(item.key);
      rows.push([
        item.part,
        item.description,
        item.groupName || 'Ungrouped',
        formulaCell(`${C}!${L.quantity}${cr}`, item.maxQuantity),
        formulaCell(`${C}!${L.salesPerYear}${cr}`, item.salesPerYear),
        formulaCell(`IF(OR(D${r}=0,${C}!${L.lineTotal}${cr}=""),"",${C}!${L.lineTotal}${cr}/D${r})`, profit?.unitPrice),
        formulaCell(`IF(D${r}=0,"",${C}!${L.landed}${cr}/D${r})`, profit?.unitCost),
        formulaCell(`IF(OR(F${r}="",G${r}=""),"",F${r}-G${r})`, profit?.unitMargin),
        formulaCell(`IF(OR(H${r}="",F${r}=0),"",H${r}/F${r})`, profit && profit.unitPrice ? profit.unitMargin / profit.unitPrice : null),
        formulaCell(`IF(F${r}="",0,E${r}*F${r})`, profit?.revenue),
        formulaCell(`IF(G${r}="",0,E${r}*G${r})`, profit?.cogs),
        formulaCell(`J${r}-K${r}`, profit?.gross),
        formulaCell(`D${r}*$B$5*12`, profit?.storage),
        formulaCell(`E${r}*($B$6+IF($B$8,0,$B$7))`, profit?.handling),
        formulaCell(`L${r}-M${r}-N${r}`, profit?.net),
      ]);
    });

    // Group summary beside the totals, driven by SUMIF over the item table.
    const groupNames = [...state.groups.map((group) => group.name), 'Ungrouped'];
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    const put = (address, cell) => { sheet[address] = typeof cell === 'object' && cell !== null ? cell : { t: typeof cell === 'number' ? 'n' : 's', v: cell }; };
    ['Sales Group', 'Items', `Revenue / yr (${o})`, 'Gross Profit / yr', 'Net Profit / yr'].forEach((label, index) => put(`${XLSX.utils.encode_col(4 + index)}12`, label));
    groupNames.forEach((name, index) => {
      const r = 13 + index;
      const summary = model.groups.find((group) => group.name === name);
      put(`E${r}`, name);
      put(`F${r}`, formulaCell(`COUNTIF(${col('C')},E${r})`, summary?.items ?? 0));
      put(`G${r}`, formulaCell(`SUMIF(${col('C')},E${r},${col('J')})`, summary?.revenue ?? 0));
      put(`H${r}`, formulaCell(`SUMIF(${col('C')},E${r},${col('L')})`, summary?.gross ?? 0));
      put(`I${r}`, formulaCell(`SUMIF(${col('C')},E${r},${col('O')})`, summary?.net ?? 0));
    });
    const lastRow = Math.max(itemEnd, 13 + groupNames.length);
    sheet['!ref'] = `A1:O${lastRow}`;
    sheet['!cols'] = [{ wch: 38 }, { wch: 28 }, { wch: 18 }, { wch: 11 }, { wch: 16 }, { wch: 18 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 13 }, { wch: 20 }, { wch: 16 }];
    styleTitle(sheet, 0, 14);
    styleHeader(sheet, 3, 0, 2);
    styleHeader(sheet, 11, 0, 1);
    styleHeader(sheet, 11, 4, 8);
    styleHeader(sheet, 24, 0, 14);
    styleGroupRow(sheet, 18, 0, 1);
    ['B5', 'B6', 'B7', 'B9', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B23'].forEach((address) => { if (sheet[address]) sheet[address].z = '#,##0.00'; });
    ['B10', 'B22'].forEach((address) => { if (sheet[address]) sheet[address].z = '0.0%'; });
    if (sheet.B21) sheet.B21.z = '0.0';
    for (let r = 13; r < 13 + groupNames.length; r += 1) ['G', 'H', 'I'].forEach((letter) => { if (sheet[`${letter}${r}`]) sheet[`${letter}${r}`].z = '#,##0'; });
    for (let r = itemStart; r <= itemEnd; r += 1) {
      ['F', 'G', 'H'].forEach((letter) => { if (sheet[`${letter}${r}`]) sheet[`${letter}${r}`].z = '#,##0.00'; });
      ['J', 'K', 'L', 'M', 'N', 'O'].forEach((letter) => { if (sheet[`${letter}${r}`]) sheet[`${letter}${r}`].z = '#,##0'; });
      if (sheet[`I${r}`]) sheet[`I${r}`].z = '0.0%';
      ['D', 'E'].forEach((letter) => { if (sheet[`${letter}${r}`]) sheet[`${letter}${r}`].z = '#,##0.##'; });
    }
    return sheet;
  }

  function buildSettingsSheet() {
    const json = JSON.stringify(collectSettings());
    const chunks = json.match(/[\s\S]{1,30000}/g) || ['{}'];
    const rows = [
      ['PartsList settings'],
      ['Machine-readable project state. Re-import this workbook into PartsList to restore mappings, variables, warehouse rates, assumptions, groups, item overrides and custom wires.'],
      ...chunks.map((chunk) => [chunk]),
    ];
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = [{ wch: 120 }];
    return sheet;
  }

  function buildProjectDataSheet() {
    const rateLabels = {
      eu120: 'EU pallet ≤120 cm / month', eu220: 'EU pallet ≤220 cm / month', sea120: 'Sea pallet ≤120 cm / month', sea220: 'Sea pallet ≤220 cm / month',
      shelf: 'Shelf location / month', drawer: 'Drawer location / month', edi: 'EDI label', receiptBase: 'Receipt base', receiptLine: 'Receipt per item line',
      orderBase: 'Order base', orderLine: 'Order per item line', parcel: 'Parcel ≤35 kg', privateSurcharge: 'Private-person surcharge', wms: 'WMS license / month',
    };
    const rows = [
      ['PartsList project data'],
      ['Readable copy of project variables, warehouse quote rates and planning inputs. Re-import the full workbook to restore the complete project state.'],
      [],
      ['Named variables', 'Value or formula', 'Source'],
      ...(state.variables.length ? state.variables.map((variable) => [variable.name, variable.expression, 'User / workbook']) : [['(none)', '', '']]),
      [],
      ['Kit routes', 'Value', 'Source'],
      ['Warehouse country', destinationConfig().name, 'Selected'],
      ['Customer country', customerConfig().name, 'Selected'],
      ['VAT treatment (customer country)', VAT_LABELS[vatTreatment()], 'Selected'],
      ['VAT rate (project setting)', state.config.vatRate, 'Selected'],
      ...Object.entries(state.kits).flatMap(([sheetName, kit]) => [
        [`${sheetName} · manufacturer`, kit.manufacturer || sheetName, 'Kit profile'],
        [`${sheetName} · origin country`, originName(kit.originCountry), 'Kit profile'],
        [`${sheetName} · source currency`, kitCurrency(kit), 'Kit profile'],
        [`${sheetName} · discount`, kitDiscount(kit), 'Kit profile'],
        [`${sheetName} · customs duty`, Number.isFinite(kit.dutyRate) ? kit.dutyRate : state.freight.dutyRate, 'Kit profile'],
        [`${sheetName} · rate buffer`, kitFxMargin(kit), 'Kit profile'],
        [`${sheetName} · freight`, kit.shippingMode === 'percent' ? `${kit.shippingAmount} % of item value` : `${kit.shippingAmount} ${kit.shippingCurrency} · ${kit.shippingMode}`, 'Kit profile'],
      ]),
      [],
      ['Warehouse quote rates', 'NOK', 'Source'],
      ...Object.entries(rateLabels).map(([key, label]) => [label, state.warehouse.rates[key], '3PL quote']),
      [],
      ['Storage plan', 'Value', 'Source'],
      ['Shelf racks with bins', state.warehouse.shelfEnabled, 'Selected'],
      ['Drawer storage', state.warehouse.drawerEnabled, 'Selected'],
      ['Pallet storage', state.warehouse.palletEnabled, 'Selected'],
      ['Planned pallet positions', state.warehouse.plannedPallets, 'Assumed'],
      ['Bins per shelf', state.warehouse.binsPerShelf, 'Assumed'],
      ['Shelf levels per rack', state.warehouse.shelvesPerRack, 'Assumed'],
      ['Units per bin', state.warehouse.unitsPerBin, 'Assumed'],
      ['Units per drawer', state.warehouse.unitsPerDrawer, 'Assumed'],
      ['Units per pallet', state.warehouse.unitsPerPallet, 'Assumed'],
      ['Pallet type', state.warehouse.palletType === 'eu' ? 'EU' : 'Sea', 'Selected'],
      ['Pallet height class (cm)', Number(state.warehouse.palletHeight), 'Selected'],
      [],
      ['Monthly activity', 'Value', 'Source'],
      ['Receipts', state.warehouse.receipts, 'Assumed'],
      ['Lines per receipt', state.warehouse.receiptLines, 'Assumed'],
      ['Orders', state.warehouse.orders, 'Assumed'],
      ['Lines per order', state.warehouse.orderLines, 'Assumed'],
      ['EDI labels', state.warehouse.ediLabels, 'Assumed'],
      ['Business parcels', state.warehouse.businessParcels, 'Assumed'],
      ['Private parcels', state.warehouse.privateParcels, 'Assumed'],
      ['Packaging / month (NOK)', state.warehouse.packaging, 'Assumed'],
      ['Include WMS license', state.warehouse.includeWms, 'Selected'],
    ];
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = [{ wch: 38 }, { wch: 24 }, { wch: 18 }];
    sheet['!freeze'] = { xSplit: 0, ySplit: 3, topLeftCell: 'A4', activePane: 'bottomLeft', state: 'frozen' };
    return sheet;
  }

  function buildWarehousingSheet() {
    const model = calculateWarehouseModel();
    const warehouse = state.warehouse;
    const rates = warehouse.rates;
    const convertedRate = (value) => model.nokToOutput === null ? null : value * model.nokToOutput;
    const rows = [
      ['Warehouse cost estimate'],
      [`Storage quantities use the imported maximum part quantities. Private rates are stored in NOK and converted to ${model.outputCurrency} with the live reference rate.`],
      [],
      ['Inventory & storage assumption', 'Value', '', 'Quoted rate', 'NOK', `${model.outputCurrency} reference`, 'Currency reference', 'Value'],
      ['Inventory units', model.inventoryUnits, '', 'EU pallet ≤120 cm / month', rates.eu120, formulaCell('IF($H$5="","",E5*$H$5)', convertedRate(rates.eu120)), 'NOK to output', model.nokToOutput],
      ['Distinct part numbers', model.distinctParts, '', 'EU pallet ≤220 cm / month', rates.eu220, formulaCell('IF($H$5="","",E6*$H$5)', convertedRate(rates.eu220)), 'Rate date', state.rateDates.NOK || ''],
      ['Shelf share', model.shelfShare, '', 'Sea pallet ≤120 cm / month', rates.sea120, formulaCell('IF($H$5="","",E7*$H$5)', convertedRate(rates.sea120)), 'Output currency', model.outputCurrency],
      ['Drawer share', model.drawerShare, '', 'Sea pallet ≤220 cm / month', rates.sea220, formulaCell('IF($H$5="","",E8*$H$5)', convertedRate(rates.sea220))],
      ['Pallet share', model.palletShare, '', 'Shelf location / month', rates.shelf, formulaCell('IF($H$5="","",E9*$H$5)', convertedRate(rates.shelf))],
      ['Bins per shelf', warehouse.binsPerShelf, '', 'Drawer location / month', rates.drawer, formulaCell('IF($H$5="","",E10*$H$5)', convertedRate(rates.drawer))],
      ['Shelf levels per rack', warehouse.shelvesPerRack, '', 'EDI label', rates.edi, formulaCell('IF($H$5="","",E11*$H$5)', convertedRate(rates.edi))],
      ['Units per bin', warehouse.unitsPerBin, '', 'Receipt base', rates.receiptBase, formulaCell('IF($H$5="","",E12*$H$5)', convertedRate(rates.receiptBase))],
      ['Units per shelf', formulaCell('B10*B12', model.unitsPerShelf), '', 'Receipt per item line', rates.receiptLine, formulaCell('IF($H$5="","",E13*$H$5)', convertedRate(rates.receiptLine))],
      ['Units per drawer', warehouse.unitsPerDrawer, '', 'Order base', rates.orderBase, formulaCell('IF($H$5="","",E14*$H$5)', convertedRate(rates.orderBase))],
      ['Units per pallet', warehouse.unitsPerPallet, '', 'Order per item line', rates.orderLine, formulaCell('IF($H$5="","",E15*$H$5)', convertedRate(rates.orderLine))],
      ['Pallet type', warehouse.palletType === 'eu' ? 'EU' : 'Sea', '', 'Parcel ≤35 kg', rates.parcel, formulaCell('IF($H$5="","",E16*$H$5)', convertedRate(rates.parcel))],
      ['Pallet height class (cm)', Number(warehouse.palletHeight), '', 'Private-person surcharge', rates.privateSurcharge, formulaCell('IF($H$5="","",E17*$H$5)', convertedRate(rates.privateSurcharge))],
      ['Selected pallet rate', formulaCell('IF(B16="EU",IF(B17=120,E5,E6),IF(B17=120,E7,E8))', model.palletRate), '', 'WMS license / month', rates.wms, formulaCell('IF($H$5="","",E18*$H$5)', convertedRate(rates.wms))],
      [],
      ['Calculated capacity', 'Count'],
      ['Shelf locations', formulaCell('IF(B5=0,0,ROUNDUP(B5*B7/B13,0))', model.shelfLocations)],
      ['Racks required', formulaCell('IF(B21=0,0,ROUNDUP(B21/B11,0))', model.racks)],
      ['Total bins', formulaCell('B21*B10', model.totalBins)],
      ['Drawer locations', formulaCell('IF(B5=0,0,ROUNDUP(B5*B8/B14,0))', model.drawerLocations)],
      ['Pallets', model.pallets],
      ['Total unit capacity', formulaCell('B21*B13+B24*B14+B25*B15', model.capacity)],
      [],
      ['Monthly operating assumption', 'Value', '', 'Monthly cost breakdown', 'NOK', model.outputCurrency],
      ['Receipts', warehouse.receipts, '', 'Shelf storage', formulaCell('B21*E9', model.costs[0].value), formulaCell('IF($H$5="","",E29*$H$5)', model.costs[0].converted)],
      ['Lines per receipt', warehouse.receiptLines, '', 'Drawer storage', formulaCell('B24*E10', model.costs[1].value), formulaCell('IF($H$5="","",E30*$H$5)', model.costs[1].converted)],
      ['Orders', warehouse.orders, '', 'Pallet storage', formulaCell('B25*B18', model.costs[2].value), formulaCell('IF($H$5="","",E31*$H$5)', model.costs[2].converted)],
      ['Lines per order', warehouse.orderLines, '', 'Receiving', formulaCell('B29*E12+B29*B30*E13', model.costs[3].value), formulaCell('IF($H$5="","",E32*$H$5)', model.costs[3].converted)],
      ['EDI labels', warehouse.ediLabels, '', 'Order handling', formulaCell('B31*E14+B31*B32*E15', model.costs[4].value), formulaCell('IF($H$5="","",E33*$H$5)', model.costs[4].converted)],
      ['Business parcels', warehouse.businessParcels, '', 'EDI labels', formulaCell('B33*E11', model.costs[5].value), formulaCell('IF($H$5="","",E34*$H$5)', model.costs[5].converted)],
      ['Private parcels', warehouse.privateParcels, '', 'Outbound freight', formulaCell('(B34+B35)*E16+B35*E17', model.costs[6].value), formulaCell('IF($H$5="","",E35*$H$5)', model.costs[6].converted)],
      ['Packaging / month', warehouse.packaging, '', 'Packaging', formulaCell('B36', model.costs[7].value), formulaCell('IF($H$5="","",E36*$H$5)', model.costs[7].converted)],
      ['Include WMS license', warehouse.includeWms, '', 'WMS license', formulaCell('IF(B37,E18,0)', model.costs[8].value), formulaCell('IF($H$5="","",E37*$H$5)', model.costs[8].converted)],
      ['', '', '', 'Monthly total', formulaCell('SUM(E29:E37)', model.monthlyNok), formulaCell('IF($H$5="","",E38*$H$5)', model.monthlyTotal)],
      ['', '', '', 'Annual total', formulaCell('E38*12', model.annualNok), formulaCell('IF($H$5="","",E39*$H$5)', model.annualTotal)],
      [],
      ['Notes'],
      ['Packaging is entered as a monthly assumption because the quote says it is charged after usage.'],
      ['API integration is excluded because the quote contains no price.'],
      ['Freight assumes one parcel per shipment up to 35 kg; private parcels add the quoted surcharge.'],
      ['WMS is excluded by default and only added when Include WMS license is TRUE.'],
    ];
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = [{ wch: 33 }, { wch: 17 }, { wch: 4 }, { wch: 32 }, { wch: 17 }, { wch: 18 }, { wch: 21 }, { wch: 18 }];
    sheet['!freeze'] = { xSplit: 0, ySplit: 4, topLeftCell: 'A5', activePane: 'bottomLeft', state: 'frozen' };
    styleTitle(sheet, 0, 7);
    [3, 19, 27].forEach((row) => styleHeader(sheet, row, 0, 7));
    [6, 7, 8].forEach((row) => { if (sheet[`B${row + 1}`]) sheet[`B${row + 1}`].z = '0.0%'; });
    for (let row = 4; row <= 38; row += 1) {
      if (sheet[`E${row + 1}`]) sheet[`E${row + 1}`].z = '#,##0.00';
      if (sheet[`F${row + 1}`]) sheet[`F${row + 1}`].z = '#,##0.00';
      if (sheet[`B${row + 1}`] && [17, 35].includes(row)) sheet[`B${row + 1}`].z = '#,##0.00';
    }
    if (sheet.H5) sheet.H5.z = '0.000000';
    styleGroupRow(sheet, 37, 3, 5);
    styleGroupRow(sheet, 38, 3, 5);
    return sheet;
  }

  function formulaCell(formula, value) {
    const cell = { t: 'n', f: formula };
    if (Number.isFinite(value)) cell.v = value;
    return cell;
  }

  function styleTitle(sheet, row, lastCol) {
    for (let col = 0; col <= lastCol; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[address]) sheet[address] = { t: 's', v: '' };
      sheet[address].s = {
        font: { name: 'Arial', sz: col === 0 ? 16 : 11, bold: col === 0, color: { rgb: '173C2B' } },
        alignment: { vertical: 'center' },
      };
    }
    sheet['!rows'] = sheet['!rows'] || [];
    sheet['!rows'][row] = { hpt: 25 };
  }

  function styleHeader(sheet, row, startCol, endCol) {
    for (let col = startCol; col <= endCol; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[address]) sheet[address] = { t: 's', v: '' };
      sheet[address].s = {
        fill: { patternType: 'solid', fgColor: { rgb: '24543D' } },
        font: { name: 'Arial', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { vertical: 'center', wrapText: true },
        border: { bottom: { style: 'thin', color: { rgb: '173C2B' } } },
      };
    }
  }

  function styleCommonRow(sheet, row, startCol, endCol) {
    for (let col = startCol; col <= endCol; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[address]) sheet[address] = { t: 's', v: '' };
      sheet[address].s = {
        ...(sheet[address].s || {}),
        fill: { patternType: 'solid', fgColor: { rgb: 'FFF2A8' } },
        font: { name: 'Arial', sz: 10, color: { rgb: '17211B' } },
      };
    }
  }

  function styleGroupRow(sheet, row, startCol, endCol) {
    for (let col = startCol; col <= endCol; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[address]) sheet[address] = { t: 's', v: '' };
      sheet[address].s = {
        fill: { patternType: 'solid', fgColor: { rgb: 'E4EFE8' } },
        font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '24543D' } },
      };
    }
  }

  function parseUserInput(text) {
    const value = String(text);
    if (value === '') return null;
    if (value.startsWith('=')) return value;
    if (value.startsWith("'=")) return value;
    if (/^(true|false)$/i.test(value)) return value.toLowerCase() === 'true';
    const number = toNumber(value);
    if (number !== null && /^[\s+\-\d.,%]+$/.test(value)) return value.trim().endsWith('%') ? number : number;
    return value;
  }

  function toNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (typeof value === 'boolean' || value === null || value === undefined || value === '') return null;
    let text = String(value).trim();
    if (!text) return null;
    const percent = text.endsWith('%');
    text = text.replace(/[^0-9,\.\-+eE]/g, '');
    if (!text || text === '-' || text === '+') return null;
    const comma = text.lastIndexOf(',');
    const dot = text.lastIndexOf('.');
    if (comma >= 0 && dot >= 0) {
      if (comma > dot) text = text.replace(/\./g, '').replace(',', '.');
      else text = text.replace(/,/g, '');
    } else if (comma >= 0) {
      const decimals = text.length - comma - 1;
      text = decimals > 0 && decimals <= 2 ? text.replace(',', '.') : text.replace(/,/g, '');
    }
    const number = Number(text);
    if (!Number.isFinite(number)) return null;
    return percent ? number / 100 : number;
  }

  function normalizeRate(value) {
    const number = toNumber(value);
    if (number === null) return 0;
    return Math.abs(number) > 1 ? number / 100 : number;
  }

  function cleanAltPart(value) {
    const text = String(value ?? '').trim();
    // "XXX.539" is a placeholder in the price list, not a number.
    return text === '0' || /^x{2,}/i.test(text) ? '' : text;
  }

  function normalizePart(value) {
    return String(value).trim().toLocaleLowerCase().replace(/\s+/g, ' ');
  }

  function normalizeCurrency(value) {
    if (value === null || value === undefined) return null;
    const text = String(value).trim().toUpperCase();
    const code = text.match(/\b[A-Z]{3}\b/)?.[0];
    return code || ({ '$': 'USD', '€': 'EUR', '£': 'GBP' }[text] ?? null);
  }

  function currencyFromText(value) {
    if (!value) return null;
    const text = String(value).toUpperCase();
    return text.match(/\b(SEK|EUR|USD|GBP|NOK|DKK|CHF|CAD|AUD|JPY|CNY|PLN)\b/)?.[1]
      || (text.includes('€') ? 'EUR' : null)
      || (text.includes('£') ? 'GBP' : null)
      || (text.includes('$') ? 'USD' : null);
  }

  const numberFormats = new Map();
  function numberFormat(options) {
    const locale = I18n ? I18n.locale(state.language) : undefined;
    const key = `${locale}|${JSON.stringify(options)}`;
    if (!numberFormats.has(key)) numberFormats.set(key, new Intl.NumberFormat(locale, options));
    return numberFormats.get(key);
  }

  function formatNumber(value, digits = 2) {
    return numberFormat({ maximumFractionDigits: digits, minimumFractionDigits: 0 }).format(value);
  }

  function formatNullableNumber(value, digits = 2) {
    return value === null || value === undefined || !Number.isFinite(value) ? '—' : formatNumber(value, digits);
  }

  function formatPercent(value) {
    return numberFormat({ style: 'percent', maximumFractionDigits: 1 }).format(value || 0);
  }

  function displayValue(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    if (isFormulaError(value)) return value.value || value.type || '#ERROR!';
    if (typeof value === 'object') return value.value ?? String(value);
    return String(value);
  }

  function isFormulaError(value) {
    return Boolean(value && typeof value === 'object' && (value.type || String(value.value || '').startsWith('#')));
  }

  function columnName(index) {
    return XLSX.utils.encode_col(index);
  }

  function populateCurrencySelect(select, selected) {
    select.replaceChildren();
    CURRENCIES.forEach((currency) => {
      const option = document.createElement('option');
      option.value = currency;
      option.textContent = currency;
      option.selected = currency === selected;
      select.append(option);
    });
  }

  function setStatus(message, type = '') {
    dom.statusMessage.textContent = message;
    dom.statusBar.classList.toggle('busy', type === 'busy');
    dom.statusBar.classList.toggle('error', type === 'error');
    dom.statusBar.classList.toggle('warning', type === 'warning');
  }

  let toastTimer;
  function showToast(message, error = false) {
    clearTimeout(toastTimer);
    dom.toast.textContent = message;
    dom.toast.classList.toggle('error', error);
    dom.toast.classList.add('show');
    toastTimer = setTimeout(() => dom.toast.classList.remove('show'), 3200);
  }

  function enabledSheetCount() {
    return Object.values(state.mappings).filter((mapping) => mapping.enabled).length;
  }

  function uniqueSheetName(base, usedNames) {
    if (!usedNames.has(base)) return base;
    let index = 2;
    while (usedNames.has(`${base} ${index}`)) index += 1;
    return `${base} ${index}`;
  }

  function quoteSheetName(name) {
    return `'${String(name).replace(/'/g, "''")}'`;
  }

  function uniqueBy(items, keyFn) {
    const seen = new Set();
    return items.filter((item) => {
      const key = keyFn(item);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }
})();
