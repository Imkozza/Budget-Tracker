# Budget Tracker — Apps Script project

A bound Google Apps Script project for a personal budget tracker: AUD, AU
financial year (1 Jul – 30 Jun), receipts stored in a dedicated Drive
folder. No external services — everything runs inside Google Sheets +
Apps Script.

This folder is the source of truth for the script. Claude Code can edit
these files and re-push them, but **the live Google Sheet + Apps Script
project themselves are not something this session can create for you** —
there's no Google Sheets/Apps Script/Drive connector available here. You
need to do a one-time deploy following the steps below (5–10 minutes).

## Build status

- [x] **Layer 1** — tabs + data model + Transactions entry form
- [x] **Layer 2** — Budget + monthly view + accountability rules
- [x] **Layer 3** — Savings goals + contributions
- [x] **Layer 4** — receipt-to-Drive upload + tax export

**The build is complete** — all four layers and every item in the original
spec's "Done when" checklist are implemented. See "What Layer 4 adds" below
and the Done-when checklist at the bottom of this file.

## What Layer 1 gives you

- All six tabs created with the exact columns from the spec: Transactions,
  Budget, Savings Goals, Savings Contributions, Reconciliation, Dashboard.
- `Transactions!Financial Year` and `Transactions!Month` auto-fill from
  `Date` via an array formula — this covers every row, whether added
  through the sidebar or typed straight into the sheet.
- Dropdown validation for Direction, Category, Scope, Reconciled (Budget
  tab), and Savings Contributions' Goal Name (pulled live from Savings
  Goals).
- **Budget Tracker → Add Transaction...** sidebar: a form to add one
  transaction at a time (Date, Amount, Direction, Category, Scope, Note,
  Receipt Link — Drive upload for it lands in Layer 4).

## What Layer 2 adds

- **Budget Tracker → Seed Budget Month...** menu item: enter a month
  (`YYYY-MM`) and it inserts one `Budget` row per budgetable category for
  that month (skipping any that already exist), `Budgeted Amount` left at
  0 for you to fill in. You can still add/edit rows by hand — the data
  model stays "one row per category per month".
- **Reconciliation** tab auto-computes `Tracked Net` (Income − Expense for
  that row's month, from Transactions) and `Difference` (`Actual Bank Net`
  − `Tracked Net`). `Difference` stays blank until you fill in `Actual Bank
  Net` for that month, so unreconciled months never show as a false
  "flagged" difference.
- **Dashboard** is now live:
  - A month picker (`B4`, pick any date in the target month) drives every
    figure below it, plus a derived Financial Year label.
  - **This month**: Income, Expenses, Net.
  - **Data quality & accountability**: Unassigned (Income − Σ budgeted
    categories for the month), Uncategorised expense count (all time),
    Work expenses missing a receipt (all time), and a count of
    Reconciliation months with a flagged Difference — every one of these
    turns red when it's non-zero.
  - **Spent vs Budgeted by category** for the selected month (one row per
    expense category, including Uncategorised, with a Remaining column
    that flags red when you've overspent).
  - **Top 5 categories by spend** for the selected month.
- **Accountability rules, enforced everywhere:**
  - *Every expense needs a Category.* Leave it blank (sidebar or typed
    directly into the sheet) and it's written back as literal
    `"Uncategorised"` and highlighted red in the Transactions grid; the
    Dashboard counts these.
  - *Zero-based check.* Dashboard's Unassigned cell flags red whenever
    Income ≠ Σ(budgeted categories) for the selected month.
  - *Reconciliation.* Enter `Actual Bank Net` per month in the
    Reconciliation tab; `Tracked Net`/`Difference` are automatic and the
    Difference cell (plus the Dashboard's flagged-months count) turns red
    on any non-zero difference.
  - *ATO substantiation.* Any Work-scope transaction with an empty
    `Receipt Link` is highlighted red in Transactions and counted on the
    Dashboard.

## What Layer 3 adds

- **Savings Goals!Current** auto-sums every matching row in Savings
  Contributions for that goal (same `ARRAYFORMULA` + criteria-range pattern
  used elsewhere), so it's always in sync — just add contribution rows.
- **Dashboard → Savings Goals Progress**: a live table, one row per goal
  (driven directly off Savings Goals, so adding a goal there just makes it
  appear here — no setup re-run needed), showing:
  - **% Complete** (`Current / Target`).
  - **Avg Monthly Contribution**: total contributed so far ÷ whole months
    since that goal's *first* contribution.
  - **Projected Completion**: today's date advanced by
    `ceil((Target − Current) / Avg Monthly Contribution)` months — or
    `"Goal reached!"` / `"No contributions yet"` where that doesn't apply.
  - **On Track?**: compares Projected Completion against the goal's
    Deadline (only shown if a Deadline is set) and flags red when
    `"Behind"`.
- Savings Contributions already had Goal Name dropdown validation (from
  Layer 1, sourced from Savings Goals) — just add rows there directly, no
  extra form needed.

## What Layer 4 adds

- **The one Drive folder** (`Budget Tracker Receipts`) — created next to
  the spreadsheet the first time you run `setupWorkbook`, or lazily on
  first use otherwise. Its ID is cached in Script Properties so it's
  reused, never recreated. **Budget Tracker → Open Receipts Folder** shows
  a clickable link to it any time.
- **Add Transaction sidebar**: a "Receipt Image" file picker (images or
  PDF, 10 MB max). Pick a file and submit — it's read client-side, sent to
  the script, uploaded into the receipts folder as
  `<date>_<scope>_<category>_<amount>.<ext>` (e.g.
  `2025-08-17_Work_Transport_42.50.jpg`), shared as "anyone with the link
  can view", and that link is written into `Receipt Link` automatically.
  The manual link field still works if you'd rather paste a link than
  upload a file (a chosen file always wins if both are filled in). If the
  upload fails for any reason, the transaction still saves — you just get
  a warning and an empty `Receipt Link` to fix up later.
- **Budget Tracker → Export Tax Year...**: enter a financial year's
  starting calendar year (e.g. `2025` for FY2025‑26). It:
  1. Filters `Transactions` for `Scope = Work` and that financial year.
  2. Writes them (Date, Amount, Category, Note, Receipt Link, sorted by
     date) to a `Tax Export` tab, with a Total row and any row missing a
     Receipt Link highlighted red.
  3. Saves a PDF snapshot of that tab into the receipts folder
     (`TaxExport_FY2025-26_<timestamp>.pdf`) and shows you both the
     summary and the PDF link in a dialog.

  The `Tax Export` tab is regenerated (not appended to) each time you run
  the export, so re-running it for the same year is safe.

### A privacy note on receipt links

Uploaded receipts are shared as **"anyone with the link can view"**, not
private-to-you-only — that's what makes the link in `Receipt Link` usable
by anyone you share the sheet or a Tax Export with (e.g. an accountant),
without them needing Drive access to your account. If you'd rather receipts
stayed fully private, open the receipts folder (menu item above) and change
its sharing in Drive; the workbook doesn't depend on the link being
public, only on it being a valid URL.

## One-time setup

1. **Create the Sheet.** Go to [sheets.google.com](https://sheets.google.com) →
   Blank spreadsheet. Name it e.g. "Budget Tracker".
2. **Open the bound Apps Script project.** Extensions → Apps Script.
3. **Bring in the code.** Easiest path — copy/paste:
   - Delete the default `Code.gs` content in the Apps Script editor.
   - For each file in this folder (`appsscript.json`, `Config.gs`,
     `Utilities.gs`, `Setup.gs`, `Code.gs`, `Transactions.gs`, `Rules.gs`,
     `Budget.gs`, `Drive.gs`, `TaxExport.gs`, `Sidebar.html`): create a
     matching file in the Apps Script editor
     (use the "+" next to Files; pick **Script** for `.gs` files and
     **HTML** for `Sidebar.html`; for `appsscript.json` you'll need
     **Project Settings → Show "appsscript.json" manifest file in editor**
     first) and paste the contents in.

   Prefer the CLI? If you have [`clasp`](https://github.com/google/clasp)
   installed and logged in:
   ```bash
   cd apps-script
   clasp create --type sheet --title "Budget Tracker" --rootDir .
   clasp push
   clasp open
   ```
   `clasp create` generates its own `appsscript.json` — overwrite it with
   the one in this folder (it sets the Queensland timezone) then `clasp push`
   again.
4. **Run setup.** In the Apps Script editor, select the `setupWorkbook`
   function (dropdown next to Run/Debug) and click **Run**. Authorize the
   requested permissions when prompted — it'll ask for the spreadsheet and
   Drive (to create/use the one receipts folder); everything stays inside
   your own Google account, there's no external service involved.
5. **Reload the spreadsheet tab.** You'll see a new **Budget Tracker** menu.
6. Use **Budget Tracker → Add Transaction...** to open the sidebar and add
   a transaction, or type directly into the Transactions tab — both are
   auto-tagged with Financial Year and Month.

Re-running `setupWorkbook` at any time is safe: it repairs headers,
formatting, and validation without touching existing data rows.

## Notes on the Financial Year formula

Australian FY: 1 Jul – 30 Jun. A transaction dated 17 Aug 2025 is FY2025‑26;
one dated 3 Feb 2026 is also FY2025‑26; one dated 1 Jul 2026 is FY2026‑27.
This is computed with a single `ARRAYFORMULA` in `Transactions!H2`/`I2` —
don't type over column H or I by hand (they're protected with a warning,
not a hard block, so you can still fix things if you ever need to).
`Reconciliation!C` (`Tracked Net`) and `!D` (`Difference`) are the same
pattern and are protected the same way, as is `Savings Goals!D` (`Current`).

## Notes on the Dashboard month picker

`Dashboard!B4` is the only cell you need to touch to change what the
Dashboard shows — pick any date inside the month you want. Everything else
(the This Month figures, the accountability flags, the category table, Top
5) reads from a hidden helper cell (`N4`, named range `DashMonth`) that
normalises your pick to the 1st of that month, so it doesn't matter which
day you choose. Columns `N`–`Q` are internal helpers (`DashMonth` plus the
savings-projection math) and are hidden by default — unhide them if you
ever want to see the working.

## Timezone

The manifest sets `Australia/Brisbane` (Queensland — no daylight saving),
so dates/timestamps in Apps Script match your local calendar.

## "Done when" checklist

From the original spec, mapped to what implements it:

- **Income/expense entries auto-tag Financial Year and Month.** ✅
  `Transactions!H2`/`I2` `ARRAYFORMULA`s (Layer 1).
- **Spent-vs-budgeted is correct per category for a chosen month.** ✅
  `Dashboard!B4` month picker + the category table at `Dashboard!A19:D37`
  (Layer 2).
- **Uncategorised, Unassigned, and Reconcile differences all flag
  visibly.** ✅ Red conditional formatting in `Transactions!D`, the
  Dashboard's Unassigned/Uncategorised/flagged-months cells, and
  `Reconciliation!D` (Layer 2).
- **Attaching a receipt lands it in Drive with a working link in the
  row.** ✅ Sidebar file picker → `uploadReceiptFile()` → `Receipt Link`
  (Layer 4).
- **Export all Work transactions for a chosen FY with receipt links + a
  total.** ✅ **Budget Tracker → Export Tax Year...** → `Tax Export` tab +
  PDF (Layer 4).

## Things worth knowing that aren't in the spec

- **Testing approach.** There's no live Google Sheet to test against from
  this session, so every layer's formulas and server-side functions
  (`setupWorkbook`, `seedBudgetMonth`, `addTransaction`,
  `uploadReceiptFile`, `exportTaxYear`) were exercised against a hand-built
  Node mock of the `SpreadsheetApp`/`DriveApp`/`UrlFetchApp` APIs to catch
  formula and logic bugs before you ever run it for real. That's not a
  substitute for actually opening the sheet and clicking through it once
  it's deployed — please do that, and say if anything looks off.
- **Category list** (`Config.gs`) is a reasonable default set for a
  personal AU budget, not something pulled from your actual spending — add
  the sidebar/menu items to edit it if you want different categories, or
  just add to the arrays in `Config.gs` and re-run setup.
- **"One row per category per month"** on `Budget` is manual entry (with
  **Seed Budget Month...** to reduce the typing) rather than
  auto-generated, since the spec's data model calls for real rows there,
  not formulas.
