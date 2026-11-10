# Practice: redo the work from the raw sheet

Ten tasks. Start from the **Raw** sheet in `excel-sales-dashboard.xlsx` (or open `data/raw_sample.csv`). Copy the data to a new workbook first so the original stays untouched. Each task ends with a number to check against; the numbers come from this workbook.

Try them with the keyboard (see [KEYBOARD.md](KEYBOARD.md)). If a check number does not match, look at the matching row in the Cleaning Log sheet.

## 1. Count and look

How many rows does Raw have? Which columns have blank cells? Count the blanks in CustomerID and in Description.
Check: 8,495 rows. Raw has 2,298 blank CustomerIDs and 23 blank Descriptions. The workbook shows 2,241 GUEST fills (R08) and 0 blank descriptions (R09) because those rules run after the removal rules R01 to R04 have taken out some of those rows.

## 2. Remove duplicates

Remove rows that are exact copies of an earlier row across all columns.
Check: 69 rows removed (R01).

## 3. Drop service lines

Some StockCodes are not products (postage, bank charges, fees, and similar). Find them by sorting or filtering the StockCode column, then delete those rows. The Cleaning Log shows which codes the workbook removed under R02.
Check: 51 rows removed.

## 4. Drop stock adjustments

Remove rows where Quantity is zero or negative and the InvoiceNo does not start with C. Keep the rows on C invoices: those are real returns.
Check: 27 rows removed (R03).

## 5. Drop bad prices

Remove rows where UnitPrice is zero or negative.
Check: 13 rows removed (R04). After tasks 2 to 5, you should have 8,335 rows. The totals are 69 + 51 + 27 + 13 = 160 removed.

## 6. Tidy text

Trim spaces in Description with `TRIM`. Rename the country EIRE to Ireland. Replace blank CustomerID with GUEST.
Check: R05 trimmed 1,994 descriptions, R07 renamed 65 rows, R08 filled 2,241 rows. These counts are on the data left after tasks 2 to 5. Raw itself has 67 EIRE rows and 2,298 blank CustomerIDs, so do this task after removing rows, not before.

## 7. Add columns

Add Type (`Return` if InvoiceNo starts with C, otherwise `Sale`), Revenue (`Quantity * UnitPrice`) and Month (year and month of InvoiceDate as text like `2011-11`).
Check: the sum of Revenue over all clean rows is 171,346.13. That is net revenue, with returns counted as negatives.

## 8. Revenue by month

Build a month table with `SUMIFS`. Which month has the highest net revenue?
Check: 2011-11 at 30,623.63. Note that 2011-12 is a partial month.

## 9. Revenue by country

Build a country table with `SUMIFS`, then work out the UK share of net revenue.
Check: the UK is 83.9% of net revenue. Germany is 1,256.50.

## 10. Count orders

Count distinct orders: one count per InvoiceNo, Sale invoices only. Do it two ways: a helper column that flags the first row of each invoice (as `OrderFlag` does), and `UNIQUE` plus `FILTER` if you have Excel 365.
Check: 302 orders overall, 4 for Germany. Net revenue divided by orders is not the same as the dashboard's average order value, which uses sales revenue (174,384.40) divided by orders: 577.43. Work out why.

## Stretch

- Build the Country dropdown and a `SUMIFS` that treats All as a wildcard (`*`).
- Rebuild the cleaning as Power Query steps using the list in the README.
- Add a chart of net revenue by month that reacts to the Country dropdown.
