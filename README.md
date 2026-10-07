# PartsList Workbook

A browser-based workbook tool for consolidating parts lists from Excel files.

## What it does

- Imports `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`, and `.csv` files locally in the browser.
- Includes a plain-language setup guide that checks sheet mappings, prices and currency rates, freight, warehouse assumptions, profitability completeness, and whether the latest changes have been exported. Each incomplete check links directly to the relevant screen.
- Shows an always-visible saved/unsaved indicator. Business data is never silently saved to browser storage; users are directed to save a re-importable Excel project workbook or password-encrypted project.
- Preserves formulas for calculation and supports linked cells across sheets.
- Detects multi-row headers and the first data row, including discount, multiplier, freight and freight-margin columns. All mappings remain editable.
- Keeps successfully detected column mappings collapsed under an advanced disclosure, while sheets requiring attention open automatically with a clear warning.
- Lists parts found on multiple enabled sheets first, highlights them in yellow, and uses the maximum quantity found.
- Lists unique parts underneath the common parts.
- Leads the consolidated table with source sheets, part, article name, a second article number where a list has one, quantity and amount in SEK.
- Makes the consolidated parts list the primary tab content: its compact overview scrolls away with the page, while the column header remains sticky. The left settings sidebar collapses to a slim rail (and expands again) from the panel's own handle, its rail, or the icon at the left of the header; the choice is remembered in the browser.
- Uses one consistent style and height for every header control, with the language switch (**SV | EN**) as a segmented control.
- Is Swedish by default and switches to English from the header. Text is translated as it is rendered, including status messages, tooltips, charts, wire diagrams, customs guidance and map labels; country names follow the language. The choice is remembered in the browser. The Swedish text lives in `i18n-sv.js` (English source string → Swedish, with `{0}` placeholders for numbers and names); `npm test` fails if a static string in the page has no translation.
- Adds kits from other manufacturers into the open project with **Add kits** (one or more files). Column names are detected in Swedish, English, German, French and Russian. Each kit has its own manufacturer, origin country, source currency, discount and customs duty; an empty discount, currency or duty means "use the project-wide value". Kits from TEI start as TEI Rock Drills / USA. An added kit takes its manufacturer from the file name, starts at zero freight, and gets its origin from an unambiguous currency in its header (for example CHF → Switzerland).
- Imports a multi-kit price list (one sheet holding many kits one after another): open the file (or use **Add kits**) and only the spare-part kits sheet is read. Every `SPARE PART KIT …` block becomes its own kit sheet (manufacturer name plus the kit name) with the manufacturer's article number, a second kit-specific article number (shown in an **Alt. part no.** column), quantity and the **Net** EUR price. These kits use the net price as given, so no further discount is applied, and start with a configurable manufacturer, currency and route (an EU origin imports into Norway as an import). The first article number identifies a part, so the same part with a different second number in another kit still counts as common; placeholder numbers such as `XXX.539` are ignored and a line with no numbers falls back to its description. Other sheets in the file are skipped, and kits that are already in the project are not added a second time. The **Alt. part no.** column is also written to the exported `Consolidated` sheet. A kit with its own discount, duty or customs-free route is exported with values in place of the shared-cell formulas, so recalculating the workbook in Excel gives the same numbers as the app.
- Is built for non-technical users. The parts list starts on a **Simple** column view (article, name, quantity, manufacturer, price per unit, line total) with one-click **Purchase**, **Sales** and **Everything** views. Every row has a **?** button that explains the price step by step (list price, exchange rate with buffer, discount, freight, margin, multiplier, line total). A **problem bar** names what is missing in plain words with a button that shows or fixes it. **Customer price list** saves an `.xlsx` with prices only. Adding a file opens a short **setup dialog** for each new manufacturer, each manufacturer card has **Reset to the project defaults**, the start screen has **Try with example data**, arrows on the map carry their freight totals, and replacing a project with unsaved changes asks first.
- Gives every manufacturer its own route. The **Manufacturers** panel in the settings (and a separate lane for each manufacturer in **Kopplingsvy / Wire view**, with its own **Shipment & tolls** table on the map) holds country of origin, price currency, discount, customs duty, an **exchange-rate buffer** (an amount added to the live rate, shown next to the live rate so you see what is used), a **kit freight** add-on (sheet values, per line, whole kit, or a **percentage of each item's value**), and whether the manufacturer **shares the project's consolidated shipment** or **ships on its own** (own shipment, clearance and insurance). Changing one manufacturer never changes another; only the warehouse is shared. Individual kit sheets keep their own overrides.
- Supports several manufacturers in one warehouse: kits from each manufacturer share one consolidated list (common parts in yellow, maximum quantity), the kit chips choose which kits count, and the list filter offers **Manufacturer: …**, **… common parts** and **… unique parts** for each manufacturer.
- Picks the warehouse country (Norway or Sweden) and treats each kit's route accordingly: goods from the warehouse's own country are domestic, goods from another EU country to a Swedish warehouse are intra-EU (no duty, no clearance fee, no import VAT), and everything else is an import with duty, clearance and import VAT. The map draws one route per manufacturer and origin.
- Sets the customer (sales) country separately from the warehouse country. It defaults to Norway; selling from a warehouse in one country to customers in another is flagged as a cross-border sale whose export documents and customer-side VAT are not calculated yet.
- Applies a configurable discount, price multiplier, freight, freight margin, and daily currency conversion.
- Replaces imported Google Finance currency formulas with the refreshed rate inside the browser calculation engine while preserving the original formulas in exported source sheets.
- Sets freight per kit: keep the sheet's `Frakt` values, apply one amount per line, or split one amount across the whole kit. Kit freight has its own currency (SEK by default, matching the sheets).
- Adds one shipment for the whole consolidated order (for example 250 USD with every kit set to 0), split across lines by purchase value, quantity, or item lines, plus customs duty, insurance, and clearance/broker fees. Freight and import costs carry the freight margin into the customer price.
- Removes or re-adds kits from the chips at the top of the consolidated list (× / +); every total, the warehouse, the dashboard, and exports recalculate.
- Calculates **without VAT by default**: the project VAT rate is 0% and locked. Unlock the padlock (sidebar, dashboard VAT box, or the customer pricing table) to enter a rate such as 25%. VAT is only ever added to the *price after margin*, never to the purchase cost, and only for sales countries set to "VAT on the price after margin". By default that is Sweden (Sweden → Sweden, Norway → Sweden where VAT is paid when you sell, and Switzerland → Sweden) and not Norway (USA → Norway and Sweden → Norway carry no VAT). The per-country setting is editable once unlocked. When VAT is not in force, the VAT columns and totals are hidden. VAT is shown separately and never counted as profit or cost.
- Sorts the consolidated parts list by any column: click a heading to sort ascending, click again to sort descending, click a third time to return to the default common-then-part order. Rows with a blank value stay at the bottom in both directions, and numbers order numerically (1 200 above 80, A2 above A10). The choice is remembered in the browser.
- Lets you choose which columns the consolidated list shows. Rarely used columns (FX rate, per-unit amounts, discount, selling price, duty and fees, pricing source) are hidden by default so the table fits a normal screen, and **Columns** in the list header shows, hides and resets them.
- Selects parts by clicking a row, by clicking its checkbox, or from the keyboard: arrow keys move, Shift+arrow extends a range, Space toggles, and Ctrl/?+A selects everything currently visible. The chosen layout and sort are remembered in the browser.
- Undoes the last 40 changes with **Undo** in the header or Ctrl/?+Z (Ctrl/?+Shift+Z redoes). Bulk actions on the selected parts count as one step, and worksheet cell and range edits restore their previous values. Ctrl/?+Z is left alone while you are typing in a field.
- Coalesces recalculation while you type, so editing a number field no longer rebuilds the whole table on every keystroke. Dropdowns and checkboxes still update at once.
- Keeps the current view in the address bar, so a view can be bookmarked or linked to and the browser Back button steps between views.
- Filters the parts list down to the rows that still need something filled in: the **Missing inputs** card in the summary row is a shortcut, and *Show ? Missing inputs* is also available directly.
- Honours a reduced-motion preference (animations and transitions are switched off) and only applies hover styling on devices that actually hover.
- Selects parts with checkboxes in the consolidated list (Shift-click for ranges, select all visible) or by row numbers in a kit tab (Ctrl/⌘ or Shift for several), then applies an action: assign to a sales group, set expected sales per year, set stock quantity, set a price multiplier, exclude/include, or clear overrides. Changes propagate to totals, warehouse capacity, the dashboard, the wire view, and exports.
- Groups parts into up to eight sales groups, each with an optional expected-sales assumption (stock turns per year or units per item) and price multiplier.
- Provides a profitability dashboard: annual net profit, revenue, gross profit, warehouse and logistics cost, payback on the landed stock investment, a profit bridge, cumulative payback curve, net profit by sales group, and per-item profitability. Each chart has a table view. A what-if comparison models changes to sales volume, purchase/FX cost, inbound freight, and warehouse/logistics without overwriting workbook inputs.
- Opens the wire view with a factory-to-customer overview: manufacturer and discount → shipment & tolls table → Norway customs → warehouse cost table → warehouse → customer pricing → customer. Values are editable in place and stay in sync with the sidebar. Per-item outbound shipment, outtake, and storage costs are calculated and locked; unlock to override them.
- Explains Norwegian customs and import VAT for the supplier's country (USA by default; EU, EFTA, and other origins differ): duty rules and EUR.1 proof of origin, import VAT reported through the MVA-melding, and ways to avoid cost and border stops (Tollager, T1 transit, Tollkreditt, DAP/FCA instead of DDP). The guide appears when you hover or edit the customs fields, can be pinned with the **i** button, and is also in the Freight & import panel.
- Switches the factory-to-customer overview between **Route** and **Map**. On the map, one arrow runs from the manufacturer across the ocean into the editable **Shipment & tolls** table and on to the destination country (kits from other origins join the same table); a second panel shows that country enlarged, with customs → warehouse and an arrow north through the editable **Customer pricing** table to the customer in the north (Tromsø in Norway, Luleå in Sweden). The **Warehouse costs** table sits by the warehouse and opens with its arrow; Customer pricing opens its per-item costs the same way. Every value stays in sync with the sidebar, and the tables are placed clear of the markers. Narrow windows stack the two panels and list the tables under the map. Hovering a wire or point shows its costs and rules. Clicking a country draws a comparison line with that origin's customs rules and can adopt it as the supplier country.
- Applies one formula or value across a row, column, or rectangular range while respecting relative and absolute references.
- Shows direct cell precedents and dependents.
- Lets users explicitly enable shelf racks, drawers, and pallets. Shelf racks with bins are enabled by default; drawer and pallet storage are opt-in, and the exact planned pallet-position count is entered directly. Bin count, units per bin, shelf levels per rack, drawer capacity, and pallet capacity are clearly marked planning assumptions.
- Calculates monthly and annual warehouse costs for storage, receipts and item lines, orders and item lines, EDI labels, outbound parcels, private-customer surcharge, packaging, and WMS after private rates are entered or re-imported.
- Leaves the optional WMS license disabled by default.
- Renders a live isometric rack, bin, drawer, and pallet preview that changes with the capacity assumptions. The preview supports drag-to-pan, wheel/button zoom, and reset controls. Its blue/orange rack and isometric construction follow the visual conventions used by [`nomsams/webware`](https://github.com/nomsams/webware), while remaining a small native SVG renderer for this app.
- Includes overview and under-the-hood wire modes for tracing color-coded relationships between source columns, calculated columns, warehouse inputs, and outputs. Mid-wire arrowheads keep direction visible when endpoint arrows sit behind nodes, outgoing branches receive distinct colors, and nodes can be dragged into a clearer layout or reset. The view can expand to the full window and accepts labeled custom relationships.
- Saves and restores the complete project as a password-encrypted `.partslist` file using browser-native AES-256-GCM. This includes the source workbook, mappings, pricing assumptions, warehouse inputs, variables, and custom wires.
- Shows private warehouse rates in their quoted NOK currency alongside live NOK→SEK reference conversions; warehouse totals are displayed in the selected output currency (SEK by default).
- **Save project .xlsx** exports one re-importable workbook with the original source sheets, formula-driven `Consolidated`, `Profitability`, `Warehousing`, and `Assumptions` sheets, a readable `Project data` sheet, and a hidden machine-readable `PartsList settings` sheet. Re-importing restores column mappings, named variables, kit freight, freight and import costs, VAT, warehouse quote rates and assumptions, storage choices, sales groups, item overrides, removed kits, wire relationships, and layout state. Generated sheets remain outside parts consolidation and re-exporting keeps stable sheet names.
- Keeps business data out of browser storage. Settings are restored from a re-imported consolidated workbook or a password-encrypted `.partslist` project. Legacy Caesar-14 project and warehouse entries are removed automatically.

## Popularity and suggested stock

The **Popularity** tab joins a purchase history (one file per market, for example Sweden and Norway for the last five years) to the kits. A history is any sheet or CSV with a column for the article number and a column for the quantity; "Item No. - description" in one column also works, and several lines for the same article are added up. Articles are matched to kit parts on the article number or the alternative number.

- **All parts** ranks every article by pieces sold, per market and in total, and also lists articles that are in no kit.
- **Popular in both** shows only articles bought in every market, ranked by the market that buys the least, so one large market cannot carry a part alone.
- **Kits** ranks the kits by the pieces sold of their parts. A part that sits in several kits counts for each of them. Click a kit to see its parts.
- **Suggested stock** suggests a quantity per part: the sales of the chosen number of months plus a safety margin for the chosen chance of having the part in stock (demand is treated as random, so the margin grows with the square root of the expected demand), and at least one piece for anything that sold. Parts are in class A, B or C by how much they sell. The suggestion can keep popular kits complete (one piece of unsold parts in kits where most parts sell) and can be fitted to a budget: safety margins go first, then slow sellers, then the rest in proportion. **Use these quantities in the project** writes the suggestion as stock quantities (it can be undone).

The consolidated list gets a **Popularity** column view and the filters *Sold in every market* and *Never sold*. The history is saved with the project and is never written to browser storage. **Save as .xlsx** on the tab exports the lists, the stock suggestion and the kit ranking.

## Saved documents (encrypted)

The project's parts lists and warehouse-rate file are not stored in readable form. `data/bundle.dat` is a single encrypted file (AES-256-GCM, key derived with PBKDF2-SHA-256, 600,000 iterations) with no readable header, file names or spreadsheet signature, so search engines and anyone browsing the repository see only random bytes. When the file is present next to the page, the start screen shows **Saved documents**; typing the key and choosing **Unlock and import** decrypts it in the browser and imports every document, rates included. Nothing is decrypted on a server and the key is never stored.

The key is not in the repository. To change the documents or the key, put the plain files somewhere outside the repository and run:

```powershell
$env:PARTSLIST_KEY = "your key phrase"
node scripts/encrypt-documents.js "parts list.xlsx" "price list.xlsx" warehouse-rates.private.json
```

Then commit the new `data/bundle.dat`. `.gitignore` blocks spreadsheets, `.partslist` files and the private rates file so a readable copy cannot be committed by mistake. Use a long key phrase: the file is public, so its strength rests on the key.

## Run locally

The project has no build step. Serve the directory with any static web server, for example:

```powershell
python -m http.server 8000
```

Then open `http://localhost:8000`.

Run the dependency-free regression tests with `npm test` or run syntax validation and tests together with `npm run check`.

The deployed version is available at https://nomsams.github.io/partslist.

## Data and rates

Workbook data stays in the browser and is not uploaded by this app. Currency rates come from the keyless Frankfurter v2 API and are cached locally for six hours. Rates are daily reference / mid-market rates, not intraday trading quotes.

The standard `.xlsx` export is readable by anyone who receives it. For access-controlled sharing, use the encrypted `.partslist` export and exchange its password separately. The password is not stored in the project file or browser storage.

Spreadsheet parsing and export use SheetJS Community Edition. Formula calculation uses HyperFormula under its GPLv3 license. Map outlines come from `world-atlas` (Natural Earth data, public domain) via jsDelivr and are drawn with `topojson-client`; they load only when the map is opened. The customs guide is general guidance, not a ruling: confirm HS codes and duty rates with your forwarder or Tolletaten.

Dashboard assumptions: expected sales come from an item override, then the item's sales group, then the default (1 × stock quantity per year). Items with no stock sell nothing unless overridden. Sold units are assumed to be replenished at the same landed cost, so payback is the landed stock investment divided by monthly net profit. Per-unit storage is monthly storage cost ÷ stock units; outtake is order handling, EDI, and packaging ÷ units sold; outbound shipment is parcel freight ÷ units sold. Receiving and WMS are fixed costs.

Private warehouse and freight figures are deliberately not embedded in the public application source. Enter them locally, re-import a previously exported consolidated workbook, or unlock an encrypted project. Packaging must be entered as a monthly estimate; API integration is excluded until a price is available. Freight assumes one parcel per shipment up to 35 kg, with the entered surcharge added for private recipients.

Warehouse quote rates can be exported or imported as a Caesar-14 encoded JSON envelope, but they are not automatically written to browser storage. A local `warehouse-rates.private.json` file is loaded when present and is excluded by `.gitignore`. Caesar-14 is reversible obfuscation rather than access control; use the AES-GCM `.partslist` format when the data must require a password.

To import a 3PL quote-rate file, open **Warehouse** and choose **Import 3PL rates JSON** in the page header or in the expanded **3PL quoted rates (NOK)** section. The importer expects a rates JSON file previously exported by PartsList; the original email text can instead be entered in the visible NOK fields. Rate JSON files do not use a password; only encrypted `.partslist` project files require the password chosen during their export.
