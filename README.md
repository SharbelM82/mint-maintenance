# Mintresponse Work Orders

A web application for tracking work orders, job profitability and receivables,
built from the **Mintresponse Work Order List** Excel tracker
(`data/work_order_template.xlsx`).

The app uses the same 37 columns (A–AK), the same drop-down lists and the same
formulas as the spreadsheet, so you can move between Excel and the app at any time:
import your filled tracker, work in the browser, and export back to Excel with
every formula intact.

## What it does

| Area | Details |
|---|---|
| **Dashboard** | The sheet's summary row (work orders, invoiced revenue, gross profit, net profit, balance due, needs financial input) plus work orders by status, receivables aging, top clients, overdue invoices and NTE exceptions. Every tile and bar links to the matching filtered list. |
| **Work orders list** | Search, filter by WO status, invoice status, client, handler, date range, "needs input", "overdue", "over NTE"; sort any column; switch between key columns and all 37 columns; running totals. |
| **Create / edit** | Form grouped like the job flows (work order → team → cost & quote → invoice → expenses → payment). Calculated fields update live while you type. Values are checked against the same rules as the Excel validation. |
| **Record a payment** | Adds to *Amount received* and sets the invoice status to *Paid* or *Partially paid* automatically. |
| **History** | Every create, edit, import and payment is logged with the fields that changed (old → new). |
| **Import** | Upload the Excel tracker (.xlsx) or a CSV with the same headings. Rows are matched on WO #: new ones are added, existing ones updated. Bad rows are skipped and listed with the reason. |
| **Export** | Excel (fills the original template, so formulas, drop-downs and formatting keep working; grows past 500 rows automatically) or CSV. Exports respect the current filters. |
| **Guide** | The definitions and worked example from the *Guide and Example* sheet, plus a column-by-column reference. |

### Formulas (identical to the Excel tracker)

| Column | Formula |
|---|---|
| J Gross profit | Invoice revenue − Direct cost |
| K Gross margin % | Gross profit ÷ Invoice revenue |
| L Expected NTE / quote | Direct cost × (1 + Target markup %) |
| X Payment due | Invoice sent date + Payment terms (days) |
| AA Actual markup % | Gross profit ÷ Direct cost |
| AE Net profit | Gross profit − Overhead − Payment/admin fees − Other expenses/tax |
| AF Net margin % | Net profit ÷ Invoice revenue |
| AH Balance due | Invoice revenue − Amount received |
| AJ NTE check | Invoice (or quote, if not invoiced) compared with the Approved NTE |
| AK Financial input status | Which input is still missing, or *Complete* |

Results stay blank until their inputs are entered, and rates show *n.a.* when the
denominator is zero, just like the spreadsheet. Amounts are rounded the way Excel
rounds. The test suite checks the app's results against the spreadsheet's own formulas.

## Running it

You need Python 3.10 or newer.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate      macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open <http://localhost:5000>. The database is created automatically in
`instance/mint.sqlite3`.

Want sample data to explore first? Run `python seed_demo.py` on an empty database.

### Sharing it on the office network

```bash
set MINT_PASSWORD=choose-a-password        # Windows (use export on macOS/Linux)
waitress-serve --host 0.0.0.0 --port 8000 app:app
```

Colleagues then open `http://<this-computer's-IP>:8000`. When `MINT_PASSWORD`
is set, everyone must sign in with that password.

### Settings (environment variables)

| Variable | Purpose |
|---|---|
| `MINT_PASSWORD` | Turns on the sign-in page with this shared password. |
| `MINT_DATABASE` | Path to the SQLite database file (default `instance/mint.sqlite3`). |
| `MINT_SECRET_KEY` | Session key. If unset, a random key is created in `instance/`. |
| `PORT`, `HOST` | Where `python app.py` listens (default `127.0.0.1:5000`). |

### Backups

All data lives in one file, `instance/mint.sqlite3`. Copy that file to back up,
or use **Export to Excel** for a readable copy.

## Project layout

```
app.py                      entry point
seed_demo.py                optional sample data
data/work_order_template.xlsx   the original Excel tracker (used for export)
mint/
  fields.py                 the 37 columns: labels, types, drop-down lists, form layout
  calc.py                   the Excel formulas in Python
  values.py                 input parsing and display formatting
  db.py                     SQLite storage and activity log
  excel.py                  Excel/CSV import and export
  views.py                  pages and actions
  templates/, static/       HTML, CSS, JS
tests/                      pytest suite
```

Run the tests with `pip install -r requirements-dev.txt && pytest`.

The earlier drafts in the repository root (`from flask import Flask, render_template.py`,
`clinet interface.py`, `mintResponse-project-latest.zip`) are left untouched and are
not used by this app.
