# 🏺 Automated_Warehouse_Engine

> **Google Apps Script · Warehouse Management System for Ceramics & Art Supplies**

A production-grade inventory tracking system built entirely in Google Sheets + Apps Script. Manages daily IN/OUT movements, cross-sheet transfers, and real-time stock levels across pottery, paints, and supplies warehouses — with zero manual formula entry.

---

## ✨ Features

| Feature | Description |
|---|---|
| **Dynamic Day Blocks** | One-click dropdown adds dated IN / OUT / Inventory columns with formulas auto-injected |
| **TOTAL NOW** | Always points to the rightmost block via `FILTER+INDEX` — survives any number of new columns |
| **Cross-Sheet Transfers** | Move quantities from an order sheet (IN or OUT) into POTTERY / PAINTS with one menu click |
| **Shelf Lookup** | Pulls shelf numbers from POTTERY storage into OUT order sheets |
| **Stock Alerts** | Conditional formatting: 🟠 orange when stock < 20, 🔴 red when stock < 10 |
| **Sort A→Z** | One-click alphabetical sort of the POTTERY sheet by description |
| **Concurrency Guard** | `LockService` prevents race conditions when two users edit simultaneously |
| **Batch API Writes** | All row updates use `setValues()` in a single call — stays within Sheets API quotas |

---

## 📁 File Structure

```
bisque-il-storage/
├── warehouse.js      # Main script: day blocks, formula engine, transfers
├── storage.js        # Supplementary: legacy transfer logic, UI helpers
└── README.md
```

### Sheet Layout

```
Spreadsheet
├── POTTERY     ← bisque items  (SKU, Description, TOTAL NOW, [daily blocks...])
├── PAINTS      ← paint supplies
├── SUPPLIES    ← consumables
├── IN          ← incoming delivery orders (source for Transfer IN)
└── OUT         ← customer orders (source for Transfer OUT + Shelf lookup)
```

---

## 🚀 Setup

### 1. Open your Google Spreadsheet

Create (or use an existing) spreadsheet with sheets named exactly:
`POTTERY`, `PAINTS`, `SUPPLIES`, `IN`, `OUT`

### 2. Add the script

1. In your spreadsheet, go to **Extensions → Apps Script**
2. Delete the default `Code.gs` content
3. Create two files: `warehouse.js` and `storage.js`
4. Paste the content of each file from this repo
5. Click **Save** (💾)

### 3. Authorize

Run any function once (e.g. `onOpen`) and follow the OAuth prompt to grant Sheets access.

### 4. Set up the trigger

In Apps Script, go to **Triggers** → **Add Trigger**:
- Function: `onEdit`
- Event type: `From spreadsheet → On edit`

The `onOpen` trigger registers automatically when you open the spreadsheet.

---

## 📋 Column Structure (per storage sheet)

```
Col A        Col B         Col C          [TOTAL NOW]   [Date block 1]          [Date block 2] ...
Description  SKU / Item#   (optional)     TOTAL NOW     IN  | Total             OUT | Total
                                          ← formula →   qty   =prev+in          qty   =prev-out
```

Row 1 = date (merged across block)
Row 2 = empty (merged, reserved for group label)
Row 3 = column headers: `IN`, `OUT`, `Total`, `TOTAL ALL` etc.
Row 4+ = data rows

---

## 🎛️ Usage

### Adding a day block

Select any storage sheet, click cell **A1**, and choose from the dropdown:

| Option | Action |
|---|---|
| `➕ IN` | Adds a new IN + Total column pair for today |
| `➕ OUT` | Adds a new OUT + Total column pair |
| `➕ Out to floor` | Adds an "out to floor" + Total pair |
| `📦 INVENTORY` | Adds STORAGE + FLOOR + TOTAL ALL inventory snapshot |
| `🔁 Update block` | Recalculates all Total and TOTAL NOW formulas |
| `🔤 Sort A→Z` | Sorts POTTERY rows alphabetically (POTTERY sheet only) |

### Transfer IN

**Transfer IN** menu → *Transfer from IN to POTTERY/PAINTS*

Reads the `IN` sheet and pushes quantities into the active IN block on the selected target sheet. Matches by SKU; appends unknown SKUs as new rows.

### Transfer OUT

**Transfer OUT** menu → *Transfer from OUT to POTTERY/PAINTS*

Reads the `OUT` sheet filtered by status column (`סטטוס pottery` or `סטטוס paints`) and pushes into the active OUT block.

**Transfer OUT** → *Shelf #*

Looks up each SKU from OUT in the POTTERY sheet and writes the shelf number back into the OUT sheet.

---

## ⚙️ Configuration

All settings are at the top of `warehouse.js`:

```javascript
const VALID_SHEETS = ["POTTERY", "PAINTS", "SUPPLIES"]; // Sheet names to manage
const DATA_START_ROW = 4;   // First data row (after 3-row header)

const ST_STORAGE_SHEET = 'POTTERY';   // Target for pottery transfers
const ST_COLORS_SHEET  = 'PAINTS';   // Target for paints transfers
const ST_ORDER_SHEET   = 'OUT';      // Source order sheet

// Column headers in the OUT sheet
const OUT_HEADER_ITEM           = 'Item #';
const OUT_HEADER_DESC           = 'Description';
const OUT_HEADER_QTY            = 'Qty';
const OUT_HEADER_STATUS_POTTERY = 'סטטוס pottery';
const OUT_HEADER_STATUS_PAINTS  = 'סטטוס paints';
```

Rename headers in the constants — no need to touch the logic.

---

## 🔢 Formula Reference

**TOTAL NOW** (auto-injected per row):
```excel
=IFERROR(
  INDEX(row:row,
    MAX(FILTER(COLUMN(F$3:Z$3), F$3:Z$3="Total"))
  ), ""
)
```
Always reads from the rightmost `Total` column — works regardless of how many blocks exist.

**Block Total (IN type)**:
```excel
=IFERROR(VALUE(TRIM(RC[-2])),0) + IFERROR(VALUE(TRIM(RC[-1])),0)
```

**Block Total (OUT type)**:
```excel
=IF(RC[-1]="", IFERROR(VALUE(TRIM(RC[-2])),0),
   IFERROR(VALUE(TRIM(RC[-2])),0) - IFERROR(VALUE(TRIM(RC[-1])),0))
```

---

## 🛡️ Error Handling

- **Block not empty**: Transfer stops and alerts the user rather than overwriting existing data
- **Wrong block type**: If the active block isn't IN when doing Transfer IN, the script aborts with a clear message
- **Missing headers**: Column finders throw descriptive errors instead of silently writing to wrong columns
- **Invalid input**: Non-numeric values in IN/OUT cells are cleared with a toast notification

---

## 🧰 Tech Stack

- **Google Apps Script** (JavaScript ES6+)
- **Google Sheets API** (`SpreadsheetApp`, `LockService`, `Utilities`)
- `onEdit` installable trigger
- `onOpen` auto-trigger for menu registration

No external libraries. No build step. Paste and run.

---

## 📸 Screenshots
[Automated_Warehouse_Engine.pdf](https://github.com/user-attachments/files/29372652/Automated_Warehouse_Engine.pdf)

> *See the portfolio presentation for visual walkthroughs of each feature.*

---

## 📄 License

MIT — free to use and adapt for your own inventory management projects.
