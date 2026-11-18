# Build flow

Order of calls from the raw CSV to the finished workbook.

Sources: [build_workbook.py](../../scripts/build_workbook.py), [cleaning.py](../../scripts/cleaning.py), [finish_in_excel.py](../../scripts/finish_in_excel.py)

```mermaid
sequenceDiagram
    participant U as User
    participant B as build_workbook.py
    participant C as cleaning.py
    participant O as openpyxl Workbook
    participant F as finish_in_excel.py
    participant X as Desktop Excel

    U->>B: python build_workbook.py stage1.xlsx
    B->>C: load_raw(data/raw_sample.csv)
    C-->>B: raw frame with raw_row
    B->>C: clean(raw)
    C-->>B: clean frame and log rows
    B->>B: add Month and Revenue columns
    B->>O: write Raw, Clean, Lists, Calc
    B->>O: write Summary, Cleaning Log, Dashboard
    B->>O: add named ranges, dropdowns, 4 charts
    B->>O: save stage1.xlsx
    U->>F: python finish_in_excel.py stage1.xlsx final.xlsx screenshots
    F->>X: open stage1.xlsx
    F->>X: add Modern Functions sheet
    F->>X: CalculateFull, scan_errors
    F->>X: export_dashboard for All, then Germany
    F->>X: SaveAs final.xlsx
```
