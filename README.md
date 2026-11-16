# AR \& AP Control Centre – Excel Portfolio Project

An Excel workbook I built to automate the accounts receivable and accounts payable work of a mid-sized German company: cleaning ERP exports, aging, KPIs, bank reconciliation, dunning, payment runs and a 13-week cash forecast. The whole cycle runs with one click.

**Tools:** Power Query · Data Model / Power Pivot · Dynamic-array formulas \& LAMBDA · VBA

> \\\\\\\*\\\\\\\*All data is fictional.\\\\\\\*\\\\\\\* \\\\\\\*RheinMain Industrietechnik GmbH\\\\\\\* and every customer, vendor, invoice and bank line were generated for this project. Base currency EUR · data Jan–Oct 2026 · as-of date 31.10.2026.

!\[Dashboard](screenshots/dashboard.png)

[<!-- Replace with your own screenshots / GIF of the one-click refresh -->](https://github.com/milanvidyapathi98-maker/AR-AP-Control-Centre/blob/main/screenshot/dashboard.png)

\---

## What it does

|Area|What I built|
|-|-|
|**Data cleaning (Power Query)**|Combines 10 messy monthly invoice exports (German format, duplicates, blank rows, `SUMME` total rows, untrimmed lower-case customer IDs) into one clean open-items table|
|**Bank import**|Parses a raw bank statement CSV and extracts invoice numbers from free-text payment references (`AR2600123`, `Re.Nr. AR26-00123`, several invoices in one line → one row per invoice)|
|**Open items \& aging**|Due dates, open amounts, FX revaluation, days past due, aging buckets and status (Paid / Partially paid / Open / Overdue / Disputed / Credit)|
|**KPIs**|Open AR, overdue %, DSO (count-back), CEI, average days late, open AP, DPO, discounts missed|
|**AP controls**|Duplicate-bill detection, bills paid twice, discounts missed, discounts taken outside the window|
|**Bank reconciliation (VBA)**|Matches every bank line to ERP receipts or vendor payments and lists exceptions|
|**Dunning (VBA)**|Builds a customer statement with dunning level 1–3 for each overdue customer and exports it to PDF|
|**Payment run (VBA)**|Selects approved, open, non-duplicate bills due within N days, applies early-payment discounts and exports a payment CSV|
|**Cash forecast**|13-week forecast from 02.11.2026 based on each customer's payment behaviour and AP due/discount dates|
|**Dashboard**|KPI cards, aging chart, AR vs AP by month, cash forecast, top-10 overdue customers; filters for Region, Currency and Sales Rep|

\---

## Highlights

### Power Query

* **Get Data → From Folder** on `Raw\\\\\\\_AR\\\\\\\_Invoices\\\\\\\\`, with German locale, cleaning, English column names and data types
* Merges with customer master data (payment terms, region, credit limit) and grouped receipts (paid amount, last payment date)
* Bank CSV import that skips the metadata lines and splits multi-invoice references into separate rows
* **Parameterised file paths:** the folder path is built from the workbook's own location, so the refresh works on any PC
* Results loaded into the **Data Model** (8 tables with relationships) for the dashboard pivots

### Formulas

* Three reusable **LAMBDA** functions in Name Manager:

  * `FXRATE(currency, date, "Avg" | "MonthEnd")`: FX lookup by month and rate type
  * `AGINGBUCKET(days)`: maps days past due to an aging bucket
  * `WEEKNO(date, start)`: assigns a date to forecast week 1–13
* `XLOOKUP`, `LET`, `IFS`, `TEXTBEFORE` / `TEXTAFTER` (parsing terms such as `2/10 Net 30`), `SUMIFS` / `COUNTIFS` / `AVERAGEIFS`
* Duplicate detection on Vendor + Amount + a *normalised* vendor invoice number (`SUBSTITUTE` / `UPPER`), flagging the later entry as the duplicate

### VBA (`vba/`)

* `RefreshEverything`: refreshes all queries synchronously, recalculates, writes a timestamp, with error handling
* `ReconcileBank`: loads tables into arrays and dictionaries for speed, matches on invoice number + date + amount within a tolerance, colours the results and writes `Recon\\\\\\\_Exceptions`
* `ExportCustomerStatements`: one PDF per overdue customer, dunning level by worst days overdue, saved in a dated folder
* `CreatePaymentRun`: InputBox for the horizon, discount logic, summary sheet + semicolon-separated payment CSV
* `RunMonthEnd`: runs all four steps in sequence
* `frmSearch`: UserForm to search a customer by ID or name and list its open items
* Columns are looked up **by header name**, not position, so the code keeps working if columns move

\---

## Results

The workbook produces these figures from the raw data:

|Check|Result|
|-|-|
|Raw invoice rows (before cleaning)|676|
|Unique AR documents after cleaning|664 (646 invoices / 18 credit notes)|
|Open AR items at 31.10.2026|179|
|Open AR (EUR, October month-end FX)|1,803,991.10|
|Open AP (EUR, October month-end FX)|232,276.42|
|Duplicate AP bills found|5 (3 already paid twice)|
|Payments with a missed discount|73|
|Discounts taken outside the window|24|
|Bank closing balance 31.10.2026|913,231.27|
|Unapplied cash (bank receipts with no ERP receipt)|6|

\---

## How to use

1. Download or clone the repository and **keep the folder structure as it is**: the workbook looks for `Raw\\\\\\\_AR\\\\\\\_Invoices\\\\\\\\` and `Raw\\\\\\\_Bank\\\\\\\\` next to itself.
2. Open `AR\\\\\\\_AP\\\\\\\_Control\\\\\\\_Centre\\\\\\\_v6.xlsm` and **enable macros and content**.
3. On the **Controls** sheet, click **Refresh Everything**, then **Reconcile Bank**, **Customer Statements** and **Create Payment Run**, or click **Month End** to run everything at once.
4. Open the **Dashboard**.

The **How\_to\_use** sheet in the workbook explains every button and sheet.

> \\\\\\\*\\\\\\\*Requirements:\\\\\\\*\\\\\\\* Microsoft 365 / Excel 2021+ on Windows (LAMBDA, XLOOKUP, TEXTBEFORE and the VBA PDF export are needed).
> If the file is opened directly from OneDrive or SharePoint, enter a local folder path in \\\\\\\*Parameters\\\\\\\* B11 and B12.

\---

## Key assumptions

* FX rates are units of foreign currency per 1 EUR; open items are revalued at the October month-end rate.
* Expected AR receipt date = due date + the customer's historical average days late. Disputed items and items more than 90 days overdue are not forecast.
* Items already past their expected date are spread by age: 1–30 days overdue → week 2, 31–60 → week 4, 61–90 → week 6.
* AP is paid on the discount date while the discount is still available, otherwise on the due date. Pending and duplicate bills are never paid.

\---

## Repository structure

```
├── AR\\\\\\\_AP\\\\\\\_Control\\\\\\\_Centre\\\\\\\_v6.xlsm   Main workbook
├── Raw\\\\\\\_AR\\\\\\\_Invoices/               10 monthly ERP invoice exports (CSV)
├── Raw\\\\\\\_Bank/                      Bank statement export (CSV)
├── vba/                           Exported VBA module and UserForm
├── screenshots/                   Dashboard and output examples
└── README.md
```

\---

## Author

**Milan Vidyapathige**, M.A. International Management \& Information Systems · B.Sc. Finance
Frankfurt am Main, Germany ·

