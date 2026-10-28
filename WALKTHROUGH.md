# Walkthrough

Each file in this repo, in plain English. Read the README first for the big picture.

## Top level

| File | What it is |
|---|---|
| `excel-sales-dashboard.xlsx` | The finished workbook. Open this one. |
| `README.md` | Landing page: what the project is, data source and licence, cleaning rules, what is live and what is pre-computed, verification, limits. |
| `WALKTHROUGH.md` | This file. |
| `KEYBOARD.md` | Keyboard shortcuts for working with the workbook. |
| `PRACTICE.md` | Ten tasks to redo by hand from the Raw sheet. |
| `LICENSE` | MIT licence for the code. The data keeps its own licence (CC BY 4.0). |
| `.gitignore` | Keeps scratch files (the `_work` folder, Excel lock files, the large source download) out of git. |

## data/

`data/raw_sample.csv` is the Raw sheet as a CSV: 8,495 rows from 400 whole invoices of the UCI Online Retail dataset, picked with random seed 42. It is the only data input to the workbook build. The full 541,909-row source file is not in the repo; download it from the UCI page linked in the README.

## scripts/

The workbook is built in three steps. Run them in this order.

1. `scripts/prepare_data.py`
   - Reads the full `Online Retail.xlsx` you downloaded.
   - Picks 400 invoices with seed 42 and writes `data/raw_sample.csv`.
   - Usage: `python scripts/prepare_data.py "path/to/Online Retail.xlsx"`

2. `scripts/build_workbook.py`
   - Reads `data/raw_sample.csv`.
   - Applies the nine cleaning rules (R01 to R09) in pandas and logs every change.
   - Writes a first workbook with openpyxl: Raw, Clean, Cleaning Log, Calc, Lists, Summary and Dashboard, with formulas, named ranges, dropdowns and charts.
   - Usage: `python scripts/build_workbook.py _work/stage1.xlsx`

3. `scripts/finish_in_excel.py`
   - Opens that first workbook in desktop Excel through pywin32.
   - Adds the Modern Functions sheet (`XLOOKUP`, `UNIQUE`, `SORT`, `FILTER`, `SORTBY`, `TAKE`). Excel has to write these itself so that they spill properly.
   - Recalculates, scans every sheet for error values, sets the dropdowns to All and then to Germany to test them, and exports the two dashboard PNGs.
   - Saves the final `.xlsx`.
   - Usage: `python scripts/finish_in_excel.py _work/stage1.xlsx excel-sales-dashboard.xlsx screenshots`

Why two stages: openpyxl can write formulas and charts but cannot write the dynamic-array formulas in a form Excel treats as spilling. Real Excel can.

## screenshots/

- `dashboard-all.png`: the Dashboard with every dropdown on All.
- `dashboard-germany.png`: the Dashboard with Country set to Germany.

Both are made by `finish_in_excel.py`: Excel prints the Dashboard sheet to PDF, and PyMuPDF turns the PDF into a PNG.

## Inside the workbook

Sheet order in the file: Dashboard, Summary, Clean, Cleaning Log, Raw, Calc, Lists, Modern Functions.

- **Dashboard**: four yellow dropdown cells (Country, Month, Customer type, Metric), six KPI cards, four charts. All numbers are `SUMIFS` formulas over named ranges.
- **Summary**: unfiltered tables by month, by country and by customer type, plus a country lookup. The lookup uses `INDEX`/`MATCH`.
- **Clean**: the cleaned rows as the table `tblClean`. Columns A to I are values; J to N (Type, Revenue, Month, CustType, OrderFlag) are formulas.
- **Cleaning Log**: one row per change, plus the rule table with counts and a check that Raw minus removed rows equals Clean.
- **Raw**: untouched sample.
- **Calc**: helper tables behind the charts. The Dashboard dropdowns feed the criteria cells at the top of this sheet.
- **Lists**: the values the dropdowns offer.
- **Modern Functions**: the Excel 365 formulas, kept apart so the rest of the workbook still works in older Excel.

Named ranges you will see in formulas: `cInvoice`, `cStock`, `cDesc`, `cQty`, `cDate`, `cPrice`, `cCust`, `cCountry`, `cType`, `cRevenue`, `cMonth`, `cCustType`, `cOrderFlag` (columns of `tblClean`); `lstCountry`, `lstMonth`, `lstCust`, `lstMetric` (dropdown lists); `SelCountry`, `SelMonth`, `SelCust`, `SelMetric` (the Dashboard dropdown cells); `CritCountry`, `CritMonth`, `CritCust` (the same choices turned into `SUMIFS` criteria, with All becoming `*`); `LookupCountry` (the Summary lookup cell).

## How a dropdown change reaches a chart

1. You pick Germany in the Country cell on the Dashboard (`SelCountry`).
2. A formula on Calc turns the choice into a criterion (`CritCountry`). All becomes the wildcard `*`.
3. `SUMIFS` formulas on Calc and on the Dashboard add up `cRevenue` and `cOrderFlag` where `cCountry` matches the criterion.
4. The charts read the Calc tables, so they redraw.

If you can explain those four steps, you can explain the whole dashboard.
