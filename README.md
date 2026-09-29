# PartsList Workbook

A browser-based workbook tool for consolidating parts lists from Excel files.

## What it does

- Imports `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`, and `.csv` files locally in the browser.
- Preserves formulas for calculation and supports linked cells across sheets.
- Detects multi-row headers and the first data row, including the workbook's `Rabatt 30%`, `Dubblering`, `Frakt`, and `Inkl Fraktmarginal` columns. All mappings remain editable.
- Lists parts found on multiple enabled sheets first, highlights them in yellow, and uses the maximum quantity found.
- Lists unique parts underneath the common parts.
- Leads the consolidated table with source sheets, part, description, quantity, amount in SEK, and amount in USD.
- Applies configurable discount (30% by default), price multiplier (2×), freight, freight margin (15%), and daily currency conversion.
- Replaces imported Google Finance currency formulas with the refreshed rate inside the browser calculation engine while preserving the original formulas in exported source sheets.
- Sets freight per kit: keep the sheet's `Frakt` values, apply one amount per line, or split one amount across the whole kit. Kit freight has its own currency (SEK by default, matching the sheets).
- Adds one shipment for the whole consolidated order (for example 250 USD with every kit set to 0), split across lines by purchase value, quantity, or item lines, plus customs duty, insurance, and clearance/broker fees. Freight and import costs carry the freight margin into the customer price.
- Removes or re-adds kits from the chips at the top of the consolidated list (× / +); every total, the warehouse, the dashboard, and exports recalculate.
- Shows prices excluding and including VAT (MVA, 25% by default) and the deductible import VAT. VAT is never counted as profit or cost.
- Selects parts with checkboxes in the consolidated list (Shift-click for ranges, select all visible) or by row numbers in a kit tab (Ctrl/⌘ or Shift for several), then applies an action: assign to a sales group, set expected sales per year, set stock quantity, set a price multiplier, exclude/include, or clear overrides. Changes propagate to totals, warehouse capacity, the dashboard, the wire view, and exports.
- Groups parts into up to eight sales groups, each with an optional expected-sales assumption (stock turns per year or units per item) and price multiplier.
- Provides a profitability dashboard: annual net profit, revenue, gross profit, warehouse and logistics cost, payback on the landed stock investment, a profit bridge, cumulative payback curve, net profit by sales group, and per-item profitability. Each chart has a table view.
- Opens the wire view with a factory-to-customer overview: manufacturer and discount → shipment & tolls table → Norway customs → warehouse cost table → warehouse → customer pricing → customer. Values are editable in place and stay in sync with the sidebar. Per-item outbound shipment, outtake, and storage costs are calculated and locked; unlock to override them.
- Explains Norwegian customs and import VAT for the supplier's country (USA by default; EU, EFTA, and other origins differ): duty rules and EUR.1 proof of origin, import VAT reported through the MVA-melding, and ways to avoid cost and border stops (Tollager, T1 transit, Tollkreditt, DAP/FCA instead of DDP). The guide appears when you hover or edit the customs fields, can be pinned with the **i** button, and is also in the Freight & import panel.
- Switches the factory-to-customer overview between **Route** and **Map**. The map draws the freight wire from the supplier to Norwegian customs, and a Norway map shows customs → warehouse → customer deliveries. Hovering a wire or point shows its costs and rules. Clicking a country draws a comparison line with that origin's customs rules and can adopt it as the supplier country.
- Applies one formula or value across a row, column, or rectangular range while respecting relative and absolute references.
- Shows direct cell precedents and dependents.
- Estimates shelf, drawer, and pallet capacity from imported maximum quantities. Bin count, units per bin, shelf levels per rack, drawer capacity, and pallet capacity are all editable assumptions.
- Calculates monthly and annual warehouse costs for storage, receipts and item lines, orders and item lines, EDI labels, outbound parcels, private-customer surcharge, packaging, and WMS after private rates are entered or re-imported.
- Leaves the optional WMS license disabled by default.
- Renders a live isometric rack, bin, drawer, and pallet preview that changes with the capacity assumptions. The preview supports drag-to-pan, wheel/button zoom, and reset controls. Its blue/orange rack and isometric construction follow the visual conventions used by [`nomsams/webware`](https://github.com/nomsams/webware), while remaining a small native SVG renderer for this app.
- Includes overview and under-the-hood wire modes for tracing color-coded relationships between source columns, calculated columns, warehouse inputs, and outputs. Mid-wire arrowheads keep direction visible when endpoint arrows sit behind nodes, and outgoing branches receive distinct colors. The view can expand to the full window and accepts labeled custom relationships.
- Saves and restores the complete project as a password-encrypted `.partslist` file using browser-native AES-256-GCM. This includes the source workbook, mappings, pricing assumptions, warehouse inputs, variables, and custom wires.
- Shows private warehouse rates in their quoted NOK currency alongside live NOK→SEK reference conversions; warehouse totals are displayed in the selected output currency (SEK by default).
- Exports a new `.xlsx` workbook with formula-driven `Consolidated`, `Profitability`, `Warehousing`, and `Assumptions` sheets, the edited source sheets, and a hidden `PartsList settings` sheet. Re-importing that export restores kit freight, freight and import costs, VAT, sales groups, item overrides, removed kits, and warehouse assumptions and rates, while keeping generated sheets out of parts consolidation. Re-exporting keeps the same sheet names.
- Remembers kit, freight, group, and sales settings per workbook name in browser storage (Caesar-14 encoded, like the rate file) and restores them when the same workbook is imported again. **Reset kit, freight & group settings** in the sidebar returns them to defaults.

## Run locally

The project has no build step. Serve the directory with any static web server, for example:

```powershell
python -m http.server 8000
```

Then open `http://localhost:8000`.

The deployed version is available at https://nomsams.github.io/partslist.

## Data and rates

Workbook data stays in the browser and is not uploaded by this app. Currency rates come from the keyless Frankfurter v2 API and are cached locally for six hours. Rates are daily reference / mid-market rates, not intraday trading quotes.

The standard `.xlsx` export is readable by anyone who receives it. For access-controlled sharing, use the encrypted `.partslist` export and exchange its password separately. The password is not stored in the project file or browser storage.

Spreadsheet parsing and export use SheetJS Community Edition. Formula calculation uses HyperFormula under its GPLv3 license. Map outlines come from `world-atlas` (Natural Earth data, public domain) via jsDelivr and are drawn with `topojson-client`; they load only when the map is opened. The customs guide is general guidance, not a ruling: confirm HS codes and duty rates with your forwarder or Tolletaten.

Dashboard assumptions: expected sales come from an item override, then the item's sales group, then the default (1 × stock quantity per year). Items with no stock sell nothing unless overridden. Sold units are assumed to be replenished at the same landed cost, so payback is the landed stock investment divided by monthly net profit. Per-unit storage is monthly storage cost ÷ stock units; outtake is order handling, EDI, and packaging ÷ units sold; outbound shipment is parcel freight ÷ units sold. Receiving and WMS are fixed costs.

Private warehouse and freight figures are deliberately not embedded in the public application source. Enter them locally, re-import a previously exported consolidated workbook, or unlock an encrypted project. Packaging must be entered as a monthly estimate; API integration is excluded until a price is available. Freight assumes one parcel per shipment up to 35 kg, with the entered surcharge added for private recipients.

Warehouse quote rates are automatically saved in browser storage and can be exported or imported as a Caesar-14 encoded JSON envelope. A local `warehouse-rates.private.json` file is also loaded when present and is excluded by `.gitignore`. Caesar-14 is reversible obfuscation rather than access control; use the AES-GCM `.partslist` format when the data must require a password.
