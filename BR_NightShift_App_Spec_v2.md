# B&R Food Services — Night Shift Reports: Full Application Spec
**Project:** Automated Night Shift Reports System
**Client:** B&R Food Services Inc.
**Stack:** n8n (automation backend) + Frontend (React or similar)
**Spec type:** Spec-Driven Development — hand this document to an AI coder as the single source of truth.

---

## 0. Document Overview

This spec covers the full end-to-end automation of the B&R Night Shift Reports workflow, which today is done manually in Excel. The workflow has four sequential parts:

| Part | Name | What it does |
|------|------|--------------|
| 1 | Get the Results | Compile orders, extract sheets, calculate shortages, add dropdowns |
| 2 | Apply the Routing | Move rows per office decisions, recalculate, drop source sheets |
| 3 | Build the Final Pick Sheets | PO report, Dry report, Freezer page, WH Pickup |
| 4 | Jetro Branch | Jetro page, PDF report, Menu page, Jetro produce sheet |

The app **guides the user through each part in order**, automating all data transformations and making the two human decision steps (Part 1 → 2 handoff, and the Driver Setup) as simple as possible — so simple that "the dumbest person" can do it without training.

---

## 1. High-Level Architecture

```
[User uploads raw Excel file]
        ↓
[Frontend: React SPA]
  - File upload
  - Progress/status per Part
  - Decision UI (dropdowns, driver ordering)
  - Preview tables
  - Download outputs
        ↓
[n8n Workflow Engine]
  - Receives files and user decisions via webhook
  - Runs Part 1, 2, 3, 4 as separate sub-workflows
  - Reads/writes Excel using a Python or Node script node
  - Returns processed files and preview data
        ↓
[Output: Excel workbook + PDF]
```

**Key principle:** n8n handles all data logic. The frontend handles UI, decisions, and file transfer only. Keep them cleanly separated.

---

## 2. Data Model — The Raw Input File

The user uploads one Excel workbook. It contains these sheets (names may vary slightly — match case-insensitively):

| Sheet name | Purpose |
|------------|---------|
| All Orders | The main order rows. Has a `Code` column. |
| item list (Web CSV) | Has SKU CODE, Category Name, Current Cost Price, Current Selling Price, Unit |
| Inventory | Has Item, Quantity On Hand, Cost, Case Avg Weight, BIN (Internal) |
| Shopping History | Has Item, Product, Bin, Unit Price, Case Price |

**Code matching rules (critical):**
- Trim whitespace from all codes before comparing.
- Treat a number and its text form as equal (e.g. `123` == `"123"`).
- Shopping History exception: a code may be missing one trailing letter (`U` or `C`). Treat codes as equal when the only difference is one trailing trailing letter on either side.

---

## 3. Part 1 — Get the Results

### 3.1 Compile the All Orders Sheet

For every row in All Orders, look up matching rows in the three source sheets by code and copy in these columns:

| Source Sheet | Source Column | Target Column in All Orders | Target Header |
|---|---|---|---|
| item list | Category Name | T | CATEGORY NAME |
| item list | Current Cost Price | U | CURRENT COST PRICE |
| item list | Current Selling Price | V | CURRENT SELLING PRICE |
| item list | Unit | AF | UNIT |
| Shopping History | Product | W | Product |
| Shopping History | Bin | X | New bin |
| Shopping History | Unit Price | Y | unit price |
| Shopping History | Case Price | Z | case price |
| Inventory | Quantity On Hand | AA | Quantity On Hand |
| Inventory | Cost | AB | Cost |
| Inventory | Case Avg Weight | AC | CASE AVG WEIGHT |
| Inventory | BIN (Internal) | AD | BIN(Internal) |

**Number formatting:** All price and weight columns use format `0.####` — up to 4 decimal places, never rounds or pads. Preserve full floating-point precision throughout.

### 3.2 Extract Jetro Source Sheet

After compilation, scan every row in All Orders:

- If the row's `Bin` value is **numeric** (including `0`) → move it to a new sheet called **Jetro source**.
- A bin of `0` counts as numeric.
- Text bins (COOLER, DRY, FREEZER, etc.) stay in All Orders.

### 3.3 Extract PO Sheet

From the rows **remaining** in All Orders after step 3.2:

- If `Quantity On Hand` is `0` → move the row to a new sheet called **PO**.

### 3.4 Calculate Shortages — Warehouse Short Sheet

Operate on what **remains in All Orders** after both extractions. Process per unique item code:

**CASE or Unit items** (where UNIT column = "CASE" or "Unit"):
```
shortage = (sum of ordered Qty for this code) - (Quantity On Hand for this code)
```

**LBS items** (where UNIT column = "LBS"):
```
shortage = (sum of ordered Qty × CASE AVG WEIGHT for this code) - Quantity On Hand
```

Rules:
- QOH counts **once per code** (not per row).
- If the result is **negative** → it is a shortage.
- Record the shortage value once, on the **first row** of the code in All Orders.
- Skip codes where Quantity On Hand is **blank** (these are instruction/fee lines).

**Build the Warehouse short sheet:**
- Contains every customer row of each shorted code.
- Columns: A = Shortages, B = UNIT (CASE/Unit/LBS), C = Update vendor (dropdown — see 3.5), then all the standard order columns.
- **Remove** the shorted rows from All Orders — Warehouse short holds the only copy.

### 3.5 Set Up Sheets and Dropdowns

**Sheet layout for All Orders, Jetro source, PO:**

| Column | Header | Content |
|--------|--------|---------|
| A | Sheet (dropdown) | Dropdown: All orders / Jetro source / PO. Default = the row's own sheet. |
| B | Vendor (route) | Dropdown: every vendor name from PO sheet + "Jetro". Default = the row's own vendor. Vendor names stored with commas removed. |
| C | Qty | On All Orders and Jetro source: mirrors the ordered Qty. On PO: Total Qty (sum per code, on first row of the code). |
| D onward | Order columns | Product Name, Code, Bin, Description, Vendor, Address, Qty, and all compiled columns. |

**Warehouse short — Update vendor dropdown (column C):**
Options = Jetro + the All Orders bin values (COOLER, COOLER-PD, DRY, FREEZER, PICK UP) + every vendor name from the PO sheet.

**Sort:** All Orders rows sorted by Product Name (A→Z).

**Driver Setup sheet:** Add an empty sheet with:
- Column A: "Order" — a dropdown of every driver name. Left blank for the office to fill in.
- Column B: "Freezer group" — auto-fills: top 3 rows = "Freezer one", rest = "Freezer two". (This can be formula-driven.)

### 3.6 End of Part 1 — Human Handoff

**At this point, the app pauses and presents the office decision UI** (see Section 6). The system cannot proceed to Part 2 until the office fills in:
1. The Sheet dropdown for each row in All Orders, Jetro source, and PO.
2. The Update vendor column for each row in Warehouse short.
3. The Driver Setup pull sequence.

---

## 4. Part 2 — Apply the Routing

### 4.1 Apply the Sheet Dropdown

For every row in All Orders, Jetro source, and PO:
- Read column A (Sheet dropdown).
- Move the row to the sheet it names.
- After moving, reset the row's Sheet tag (column A) to match its new home sheet.
- Rows already on the correct sheet stay put.

### 4.2 Apply Warehouse Short Routing

For each row in Warehouse short, read the Update vendor column (C):

| Office choice | Action |
|--------------|--------|
| "Jetro" | Move row to Jetro source. Remove from Warehouse short. |
| A vendor name | Move row to PO. Remove from Warehouse short. |
| "WH" | Move row to PO, set vendor = "WH". Remove from Warehouse short. |
| A bin (COOLER, COOLER-PD, DRY, FREEZER, PICK UP) | Move row back to All Orders. Overwrite Bin column with the chosen bin. Remove from Warehouse short. |

**Shortage value re-homing:** When a row carrying the shortage value (column A of Warehouse short) is moved out, re-assign the shortage value to the **first remaining row** of the same code in Warehouse short.

### 4.3 Recalculate PO Total Qty

After all moves: recalculate column C of the PO sheet — sum of ordered Qty per item code, shown on the first row of each code. Clear the value on non-first rows.

### 4.4 Confirm Driver Pull Sequence

The Driver Setup sheet is now filled. Read and store the pull sequence — it feeds into Part 3 and Part 4. Do not alter it.

### 4.5 Drop Source Sheets

Remove these sheets from the workbook: Inventory, item list (Web CSV), Shopping History. They are no longer needed.

---

## 5. Part 3 — Build the Final Pick Sheets

### 5.1 Re-order the All Orders Sheet

Sort All Orders so rows follow the driver pull sequence from the Driver Setup sheet. First driver's rows first, last driver's rows last. This is the master order.

### 5.2 Block 1 — PO by Vendors

Source: PO sheet. Output: a report grouped by vendor.

- Sorted: by vendor, then by item code within each vendor.
- Each vendor group has a spanned header: Vendor name + date. The header text is **centered**.
- **Page break rule (keep-together):** Each vendor group is an atomic block — it must never be split across pages. Before placing a group, calculate its full rendered height. If it fits in the remaining space on the current page, place it there. If it does not fit, start a new page for it. Font size is the only lever: reduce it until the group fits on one page. Do not split a vendor group across pages under any circumstances.

| Col | Header | Source |
|-----|--------|--------|
| A | Qty | Ordered Qty |
| B | total QTY | Sum of Qty per code, on first row of code |
| C | Product Name | Product Name |
| D | Code | Item code |
| E | Price | Price (selling price) |
| F | CURRENT COST PRICE | Current cost price |
| G | Jetro cost | Case price |
| H | Name | Customer name |
| I | Driver | Driver |

### 5.3 Block 2 — Dry by Driver

Source: All Orders rows where Bin = "DRY".

- Grouped by driver, each group has a spanned header: Driver name + date. The header text is **centered**.
- Driver groups ordered by the pull sequence.
- Within each driver group, rows sorted by Internal bin (BIN(Internal) column) so the picker walks the area in location order.
- **Page break rule (keep-together):** Each driver group is an atomic block — it must never be split across pages. Before placing a group, calculate its full rendered height. If it fits in the remaining space on the current page, place it there. If it does not fit, start a new page for it. Font size is the only lever: reduce it until the group fits on one page. Do not split a driver group across pages under any circumstances.

| Col | Header | Source |
|-----|--------|--------|
| A | Internal bin | BIN(Internal) |
| B | Bin | Bin ("DRY") |
| C | QTY | Ordered Qty |
| D | Product Name | Product Name + Description merged into one cell |
| E | Vendor | Vendor |
| F | Customer | Customer name |
| G | Driver | Driver |
| H | QTY | Quantity On Hand |

### 5.4 Block 3 — Freezer Page

Source: All Orders rows where Bin = "FREEZER".

- Split into two groups:
  - **Freezer one:** first 3 drivers in the pull sequence.
  - **Freezer two:** all other drivers.
- Each group has a spanned header: date + group name. The header text is **centered**.
- Within each freezer group, rows are sorted by Product Name (A→Z) so repeated codes fall on adjacent rows.
- **Page break rule (keep-together):** Each freezer group (Freezer one / Freezer two) is an atomic block — it must never be split across pages. Before placing a group, calculate its full rendered height. If it fits in the remaining space on the current page, place it there. If it does not fit, start a new page for it. Font size is the only lever: reduce it until the group fits on one page. Do not split a freezer group across pages under any circumstances.

| Col | Header | Source / Rule |
|-----|--------|--------------|
| A | Internal Bin | BIN(Internal) |
| B | qty | Ordered Qty |
| C | Total Qty | Sum of ordered Qty per item code within the freezer group, written on the first row of the code; blank on subsequent rows of the same code |
| D | product name | Product Name |
| E | Customer | Customer name |
| F | Qty on Hand | Quantity On Hand |

### 5.5 Block 4 — WH Pickup

Source: All Orders rows — all bins (COOLER, COOLER-PD, DRY, FREEZER, PICK UP, and any others).

- One block per **customer**, with a spanned header: Customer name + Driver name + date. The header text is **centered**.
- Customer blocks ordered by driver pull sequence.
- Within each customer block, rows sorted by bin in this order: Cooler → Cooler-pd → Dry → Freezer → any other bin → then by Internal bin within each bin type.
- **Page break rule (keep-together):** Each customer block is an atomic block — it must never be split across pages. Before placing a block, calculate its full rendered height. If it fits in the remaining space on the current page, place it there. If it does not fit, start a new page for it. Font size is the only lever: reduce it until the block fits on one page. Do not split a customer block across pages under any circumstances.

| Col | Header | Source |
|-----|--------|--------|
| A | qty | Ordered Qty |
| B | Product Name | Product Name + Description merged into one cell |
| C | bin | Bin |
| D | internal bin | BIN(Internal) |
| E | vendor | Vendor |
| F | Driver | Driver |
| G | qty | Quantity On Hand |

---

## 6. Part 4 — The Jetro Branch

### 6.1 Input

The Jetro source sheet from Part 2. It carries a 35-column layout. The relevant columns for Part 4 are:

| Column | Header | Used for |
|--------|--------|---------|
| C | Qty | Ordered quantity |
| D | Product Name | Part of product description merge |
| E | Code | Item code — appended to merge |
| F | Bin | Numeric Jetro bin |
| G | Description | Part of product description merge |
| M | Price | Sell price (Jetro produce) |
| O | Name | Customer name — drives grouping and sort |
| R | Transaction Date | Delivery date in block headers |
| S | Driver | Driver name — drives pull sequence |
| W | CATEGORY NAME | Identifies Produce rows |
| X | CURRENT COST PRICE | Cost price (Jetro produce) |
| Z | Product | Part of product description merge |
| AA | New bin | Alternate bin (shown only when different from Bin) |
| AB | unit price | JTR U COST (Jetro produce) |
| AC | case price | JTR C cost (Jetro produce) |
| AD | Quantity On Hand | QTY column on outputs |

### 6.2 Driver Pull Sequence

The Jetro branch uses the same driver pull sequence as Part 3, sourced from the Driver Setup sheet:

Default built-in order (from `_drivers` helper sheet if present):
```
1: Glen  2: Abraham  3: Ramon  4: Gilberto  5: Jose  6: Adrian  7: Luis  8: Z
```
Any driver not in this list sorts after Z, alphabetically.

### 6.3 Sort Jetro Source

Sort the Jetro source sheet on three keys in this priority:
1. Driver — by pull sequence (Glen first, unlisted last).
2. Customer name (column O) — A→Z within each driver.
3. Bin (column F) — ascending within each customer.

This sort is done **once** and all subsequent outputs inherit it.

### 6.4 Jetro Page

One block per customer order. Structure:

- **Header row** spanning all 5 columns: Customer name | Delivery date | Driver
- **Column header row**
- **Order line rows**

| Col | Header | Source / Rule |
|-----|--------|--------------|
| A | Bin | Column F |
| B | New bin | Column AA. **Left blank when it equals Bin.** |
| C | Qty | Column C (ordered quantity). Centered. |
| D | Product | Merge: Product Name + Description + Product + Code, joined with `/` no spaces: `ProductName/Description/Product/Code` |
| E | QTY | Column AD (Quantity On Hand). Centered. |

**Produce rows:** Any row where CATEGORY NAME (W) = "Produce" is shaded light grey.

### 6.5 Jetro Report PDF

Render the Jetro page as a printable landscape PDF:

- Landscape orientation; page number on every page.
- White column headers, black text, black borders throughout.
- Product text wraps onto a second line — never clipped.
- **Page break rule (keep-together, absolute):** Never split a customer order across pages under any circumstances. Each customer block is an atomic unit. Before placing a block, calculate its full rendered height. If it fits in the remaining space on the current page, place it there. If it does not fit, start a new page for that block. Font size is the only lever — shrink the font until the entire block fits on a single page. There is no minimum font floor that overrides this rule; fitting on one page always wins.
- **No "(continued)" bands.** Since overflow is never permitted, continuation headers are not needed and must not be generated.
- Block headers (Customer name | Delivery date | Driver) are **centered** across all columns.
- Small orders may share a page — no blank pages.
- Adaptive font: start larger on pages with lighter loads, reduce only as needed to keep each block whole on one page.

### 6.6 Menu Page

Built **after** the PDF, by reading back the actual page numbers from the finished PDF.

| Col | Header | Content |
|-----|--------|---------|
| A | Page number | PDF page number. Multi-page orders show a range, e.g. `3-6`. |
| B | Customer Orders | Every customer order header on that page, stacked in one cell. |
| C | Pulled by | Left blank — warehouse writes in who pulled it. |

**Important:** Page numbers must be read from the rendered PDF, not estimated from row counts. Row-count estimates drift as soon as text wraps.

### 6.7 Jetro Produce Sheet

A separate sheet containing only the rows where CATEGORY NAME = "Produce". Sorted by Bin, then by Product Name within the same Bin. Where the same product name repeats, sum its Qty into the **Total qty** cell of the first occurrence; leave Total qty blank on later occurrences.

| Col | Header | Source / Rule |
|-----|--------|--------------|
| A | Bin | Column F |
| B | New bin | Column AA. Blank when equal to Bin. |
| C | Qty | Column C (ordered quantity) |
| D | Total qty | Sum of Qty for repeated product name, on first row |
| E | Product Info | Merge: Product Name + Product + Code, slash-separated, no spaces: `ProductName/Product/Code` |
| F | Sell price | Column M |
| G | Cost price | Column X (CURRENT COST PRICE) |
| H | JTR U COST | Column AB (unit price) |
| I | JTR C cost | Column AC (case price) |
| J | Customer name | Column O |
| K | Driver name | Column S |

Price columns (F–I) use format `0.####`.

---

## 7. Frontend UX Specification

### 7.1 Overall Flow (Wizard-style)

The app is a **step-by-step wizard**. Users cannot skip steps. The UI makes it crystal clear what step they are on and what action is required.

```
Step 1: Upload File
Step 2: Review & Confirm (Part 1 output preview)
Step 3: Make Routing Decisions (human decisions for Part 2)
Step 4: Apply Routing (automated)
Step 5: Review & Download Final Reports
```

### 7.2 Step 1 — Upload

- Large drag-and-drop zone: "Drop your All Orders Excel file here, or click to browse."
- Accept `.xlsx` only.
- Show file name and size after selection.
- Single "Process File" button triggers n8n Part 1 workflow.
- Show a progress indicator while processing (Part 1 takes a few seconds).
- On completion, move to Step 2.
- On error: show a clear error message with the specific problem (e.g. "Could not find 'Inventory' sheet in the uploaded file").

### 7.3 Step 2 — Review (Part 1 Output)

Show a summary card for each extracted sheet:

- **All Orders:** Row count, list of unique product names (scrollable).
- **Jetro source:** Row count.
- **PO:** Row count.
- **Warehouse short:** Row count + list of shorted item codes and shortage amounts.
- **Driver Setup:** Empty — user will fill in Step 3.

Include a "Download Part 1 Workbook" button so the office can open the raw file if they want to inspect it.

Then a prominent "Proceed to Routing Decisions →" button.

### 7.4 Step 3 — Routing Decisions (Human Input)

This is the most important UX step. It has three sub-sections:

#### 7.4.1 Sheet Routing Table

Display a table with all rows from All Orders, Jetro source, and PO. Each row shows:
- Product Name
- Current sheet
- Code
- Bin
- Vendor
- **Sheet dropdown** (editable): All orders / Jetro source / PO
- **Vendor (route) dropdown** (editable): all vendor names + Jetro

Make the table filterable by current sheet and searchable by product name.

Include a "Mark all as default" button that sets all dropdowns to their current sheet (no change).

#### 7.4.2 Warehouse Short Routing Table

Display all Warehouse short rows. Each row shows:
- Item code
- UNIT type
- Shortage value
- Product name
- All order columns
- **Update vendor dropdown** (editable): Jetro / bin names / all vendor names from PO

#### 7.4.3 Driver Setup

A drag-and-drop list of all driver names found in the data. User drags them into the order they leave the warehouse.

- List defaults to: Glen, Abraham, Ramon, Gilberto, Jose, Adrian, Luis, Z (if those drivers exist in the data), with any others at the bottom.
- Auto-labels: top 3 = "Freezer one", rest = "Freezer two" — updates live as the user drags.
- Show a clear visual separator between the Freezer one and Freezer two groups.

After filling all three sub-sections, a "Apply Routing →" button triggers the n8n Part 2 workflow.

### 7.5 Step 4 — Apply Routing (Automated)

- Show a processing screen: "Applying your routing decisions…"
- Progress bar with sub-steps: Moving rows → Recalculating PO → Building reports → Generating PDFs (PO, Dry, Freezer, WH Pickup, Jetro).
- On completion, show a summary: how many rows moved where.
- Move to Step 5.

### 7.6 Step 5 — Download Final Reports

Show four download cards. Each card offers **both an Excel download and a PDF download button**:

| Card | Excel contents | PDF contents |
|------|---------------|-------------|
| PO Report | PO by vendors | Same layout, print-ready PDF |
| Dry + Freezer + WH Pickup | All three pick sheets | Same layout, print-ready PDF |
| Jetro Pack | Jetro page + Menu page + Jetro produce | — |
| Jetro PDF | — | The printable landscape Jetro report PDF |

Include a "Download All" button that zips all Excel and PDF files together.

Also show a "Start a New Night Shift" button that resets to Step 1.

### 7.7 Error Handling (General)

Every error message must be in plain English with no technical jargon. Examples:
- "We couldn't find the Shopping History sheet in your file. Please check the file and try again."
- "Row 47 (Chicken Breast) couldn't be matched to any item in the item list."
- "The routing was applied, but 3 rows in Warehouse short had no Update vendor selected. Please go back and fill those in."

---

## 8. n8n Workflow Architecture

### 8.1 Webhooks / Endpoints

| Endpoint | Trigger | Action |
|----------|---------|--------|
| `POST /part1` | File upload | Run Part 1 processing, return preview data + workbook |
| `POST /part2` | User decisions JSON | Run Part 2 routing, return updated workbook |
| `POST /part3-4` | Trigger after Part 2 | Run Part 3 + Part 4, return all output files |
| `GET /status/:jobId` | Polling | Return job status and progress |

### 8.2 n8n Sub-Workflow Structure

Recommended n8n layout:

```
Main Workflow
├── Webhook: Receive file + trigger Part 1
├── Sub-workflow: Part 1
│   ├── Parse Excel (Code node / Python)
│   ├── Compile All Orders
│   ├── Extract Jetro source
│   ├── Extract PO
│   ├── Calculate Shortages → Warehouse short
│   ├── Set up dropdowns
│   └── Return workbook + preview JSON
├── Webhook: Receive decisions + trigger Part 2
├── Sub-workflow: Part 2
│   ├── Apply Sheet dropdowns
│   ├── Apply Warehouse short routing
│   ├── Recalculate PO Total Qty
│   └── Drop source sheets
├── Sub-workflow: Part 3
│   ├── Re-order All Orders
│   ├── Build PO by vendors (Excel)
│   ├── Render PO PDF (keep-together: one vendor group per page)
│   ├── Build Dry by driver (Excel)
│   ├── Render Dry PDF (keep-together: one driver group per page)
│   ├── Build Freezer page (Excel)
│   ├── Render Freezer PDF (keep-together: one freezer group per page)
│   ├── Build WH Pickup (Excel)
│   └── Render WH Pickup PDF (keep-together: one customer block per page)
└── Sub-workflow: Part 4
    ├── Sort Jetro source
    ├── Build Jetro page (Excel)
    ├── Render Jetro PDF (keep-together: one customer order per page, centered headers)
    ├── Read PDF page numbers → Build Menu page
    └── Build Jetro produce sheet
```

### 8.3 Excel Processing

Use a Python Code node in n8n with `openpyxl` for reading and writing Excel. Key rules:

- All code matching: normalize to string, strip whitespace, compare lowercase.
- Shopping History code fuzzy match: strip trailing `U` or `C` before comparing.
- Number format `0.####`: apply `openpyxl` number format string `'0.####'` to all price/weight cells.
- Dropdown lists: use `openpyxl` `DataValidation` with `type="list"` for all dropdowns.

### 8.4 PDF Generation

**All four Part 3 reports and the Jetro Report (Part 4) are rendered as PDFs.** Use Python `WeasyPrint` (preferred) or `ReportLab` for all PDF generation.

**Universal keep-together algorithm (applies to every PDF):**

The same rendering logic is used across all five PDFs. The atomic unit differs per report:

| Report | Atomic unit (never split) |
|--------|--------------------------|
| PO by vendors | One vendor group |
| Dry by driver | One driver group |
| Freezer page | One freezer group (Freezer one / Freezer two) |
| WH Pickup | One customer block |
| Jetro Report | One customer order |

Algorithm for each atomic unit:
1. Render the block off-screen to measure its height.
2. Compare against remaining vertical space on the current page.
3. If it fits → place it on the current page.
4. If it does not fit → start a new page, then place it.
5. Font size is the only lever for oversized blocks: binary-search downward until the block fits on one page. There is no hard minimum font size — legibility is secondary to the no-split rule.

**Header centering:** All spanned block headers (vendor name + date, driver name + date, customer name + driver + date) must be **horizontally centered** within their spanning cell.

**For the Jetro Report PDF specifically:**
- After rendering, read back the actual page number of each customer block.
- Return both the PDF file and a page-number map (JSON) to build the Menu page.
- No "(continued)" bands — overflow is never permitted, so they are never needed.

---

## 9. Data Contracts (Frontend ↔ n8n)

### 9.1 Part 1 Response (n8n → Frontend)

```json
{
  "jobId": "abc123",
  "status": "complete",
  "preview": {
    "allOrders": { "rowCount": 142, "productNames": ["Chicken Breast", "..."] },
    "jetroSource": { "rowCount": 38 },
    "po": { "rowCount": 24 },
    "warehouseShort": {
      "rowCount": 7,
      "shortages": [
        { "code": "BF1234", "productName": "Beef Patty", "unit": "CASE", "shortage": -3 }
      ]
    },
    "drivers": ["Glen", "Abraham", "Ramon", "Gilberto", "Jose"],
    "vendors": ["Sysco", "US Foods", "Restaurant Depot"]
  },
  "workbookUrl": "/files/abc123/part1.xlsx"
}
```

### 9.2 Part 2 Request (Frontend → n8n)

```json
{
  "jobId": "abc123",
  "sheetRoutingDecisions": [
    { "rowId": "row_001", "sheet": "Jetro source", "vendor": "Jetro" },
    { "rowId": "row_002", "sheet": "All orders", "vendor": "Sysco" }
  ],
  "warehouseShortDecisions": [
    { "rowId": "ws_001", "updateVendor": "DRY" },
    { "rowId": "ws_002", "updateVendor": "Jetro" }
  ],
  "driverPullSequence": ["Glen", "Ramon", "Abraham", "Gilberto", "Jose", "Adrian"]
}
```

### 9.3 Final Output Response (n8n → Frontend)

```json
{
  "jobId": "abc123",
  "status": "complete",
  "outputs": {
    "poReport":      { "xlsx": "/files/abc123/po_report.xlsx",    "pdf": "/files/abc123/po_report.pdf" },
    "drySheet":      { "xlsx": "/files/abc123/dry.xlsx",          "pdf": "/files/abc123/dry.pdf" },
    "freezerSheet":  { "xlsx": "/files/abc123/freezer.xlsx",      "pdf": "/files/abc123/freezer.pdf" },
    "whPickup":      { "xlsx": "/files/abc123/wh_pickup.xlsx",    "pdf": "/files/abc123/wh_pickup.pdf" },
    "jetroWorkbook": { "xlsx": "/files/abc123/jetro.xlsx" },
    "jetroPdf":      { "pdf": "/files/abc123/jetro_report.pdf" }
  }
}
```

---

## 10. Business Rules — Quick Reference (Canonical)

| Rule | Detail |
|------|--------|
| Code matching | Trim whitespace, normalize to string, case-insensitive |
| Shopping History code | Trailing U or C may be absent — still treat as match |
| Number format | `0.####` — up to 4 decimals, never rounds or pads |
| Jetro rows | Any row where Bin is numeric (including 0) |
| PO rows | Any row where Quantity On Hand = 0 (after Jetro extraction) |
| Shortage (CASE/Unit) | sum(ordered Qty) - QOH |
| Shortage (LBS) | sum(ordered Qty × CASE AVG WEIGHT) - QOH |
| QOH per code | Counted once per code, not per row |
| Shortage record | Written on first row of the code only |
| Shortage re-home | When carrier row moves out, value goes to next remaining row of same code |
| Vendor names in dropdowns | Commas removed |
| Freezer one | First 3 drivers in pull sequence |
| Freezer two | All other drivers |
| Driver pull sequence | Single source of truth: Driver Setup sheet |
| Jetro driver order | Glen, Abraham, Ramon, Gilberto, Jose, Adrian, Luis, Z; unlisted alphabetical last |
| New bin in Jetro | Blank when equal to Bin; only show when different |
| Produce rows | CATEGORY NAME = "Produce" — shaded grey on Jetro page, also on Jetro produce sheet |
| Product merge (Jetro page) | `ProductName/Description/Product/Code` — slash, no spaces |
| Product merge (Jetro produce) | `ProductName/Product/Code` — slash, no spaces |
| Dry + WH Pickup Product merge | `ProductName + Description` — single cell |
| PO Total Qty | Sum per code, first row of code only |
| Freezer Total Qty | Sum per code within the freezer group, first row of code only; blank on subsequent rows |
| Freezer page sort | Sorted by Product Name (A→Z) within each freezer group so repeated codes are adjacent |
| WH Pickup bin order | Cooler → Cooler-pd → Dry → Freezer → others → then by internal bin |
| Menu page numbers | Read from rendered PDF — never estimated |
| Source sheets dropped in Part 2 | Inventory, item list (Web CSV), Shopping History |
| Keep-together rule | Every atomic block (vendor group / driver group / freezer group / customer block / customer order) must fit on a single page — never split across pages |
| Font sizing for keep-together | Font size is the only lever for oversized blocks; shrink until the block fits — no hard minimum |
| Block header alignment | All spanned headers (vendor + date, driver + date, customer + driver + date) are horizontally centered |
| PDF outputs | Every report has a PDF counterpart: PO, Dry, Freezer, WH Pickup, and Jetro |
| No continuation bands | "(continued)" carry-over headers are never generated — overflow is not permitted |

---

## 11. Non-Functional Requirements

- **Security:** Files are processed server-side and stored temporarily. Delete all uploaded and generated files after 24 hours. Do not store any business data long-term.
- **Performance:** Part 1 should complete in under 30 seconds for a typical workbook (< 500 rows). PDF generation may take up to 60 seconds.
- **Reliability:** n8n workflows must return structured errors, not silent failures. Every error must surface in the frontend in plain English.
- **Accessibility:** The frontend must be usable on a desktop browser. Mobile support is not required.
- **State persistence:** The job ID persists in the browser session. If the user closes and reopens the tab, they can re-enter the job ID to resume from their last step.

---

## 12. Suggested Tech Stack Details

| Layer | Recommended | Notes |
|-------|-------------|-------|
| Frontend | React + Vite + TailwindCSS | Simple and fast. Use shadcn/ui for table and drag-drop components. |
| Drag-and-drop driver list | `@dnd-kit/core` | Smooth drag-and-drop for the driver sequence. |
| Table with dropdowns | TanStack Table | Handles large row counts with virtual scroll. |
| n8n | Self-hosted or n8n Cloud | Use Code nodes (Python) for Excel processing. |
| Excel processing | Python `openpyxl` | In a Python Code node in n8n. |
| PDF generation | Python `WeasyPrint` or `ReportLab` | In a Python Code node. WeasyPrint preferred for layout flexibility. |
| File storage (temp) | n8n local filesystem or S3 | S3 recommended for production. |
| Job status polling | Frontend polls `GET /status/:jobId` every 3 seconds | Or use n8n's webhook response with a long timeout. |

---

## 13. Out of Scope (for v1)

- Multi-user / authentication (single-user tool for now)
- Editing compiled data in the browser (downloads only)
- Historical report storage
- Mobile support
- Any output format other than Excel and PDF

---

*End of Spec — B&R Food Services Night Shift Reports Automation v1.0*
