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
- [ ] Layer 2 — Budget + monthly view + accountability rules
- [ ] Layer 3 — Savings goals + contributions
- [ ] Layer 4 — receipt-to-Drive upload + tax export

## What Layer 1 gives you

- All six tabs created with the exact columns from the spec: Transactions,
  Budget, Savings Goals, Savings Contributions, Reconciliation, Dashboard
  (Dashboard is a placeholder until Layer 2).
- `Transactions!Financial Year` and `Transactions!Month` auto-fill from
  `Date` via an array formula — this covers every row, whether added
  through the sidebar or typed straight into the sheet.
- Dropdown validation for Direction, Category, Scope, Reconciled (Budget
  tab), and Savings Contributions' Goal Name (pulled live from Savings
  Goals).
- **Budget Tracker → Add Transaction...** sidebar: a form to add one
  transaction at a time (Date, Amount, Direction, Category, Scope, Note,
  Receipt Link as a manual paste-in for now — Drive upload lands in Layer 4).

## One-time setup

1. **Create the Sheet.** Go to [sheets.google.com](https://sheets.google.com) →
   Blank spreadsheet. Name it e.g. "Budget Tracker".
2. **Open the bound Apps Script project.** Extensions → Apps Script.
3. **Bring in the code.** Easiest path — copy/paste:
   - Delete the default `Code.gs` content in the Apps Script editor.
   - For each file in this folder (`appsscript.json`, `Config.gs`,
     `Utilities.gs`, `Setup.gs`, `Code.gs`, `Transactions.gs`,
     `Sidebar.html`): create a matching file in the Apps Script editor
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
   requested permissions when prompted (this project only touches the
   spreadsheet and, from Layer 4, its own Drive receipts folder — it never
   talks to anything outside your Google account).
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

## Timezone

The manifest sets `Australia/Brisbane` (Queensland — no daylight saving),
so dates/timestamps in Apps Script match your local calendar.
