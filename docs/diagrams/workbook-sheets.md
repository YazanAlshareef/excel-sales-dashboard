# Workbook sheets

How the sheets relate. Links are by named range or by a column value.

Sources: [build_workbook.py](../../scripts/build_workbook.py), [finish_in_excel.py](../../scripts/finish_in_excel.py), [excel-sales-dashboard.xlsx](../../excel-sales-dashboard.xlsx)

```mermaid
erDiagram
    RAW ||--o{ CLEAN : "CLEAN.RawRow is the Raw row number"
    RAW ||--o{ CLEANING_LOG : "CLEANING_LOG.RawRow"
    CLEAN ||--o{ CALC : "SUMIFS over c* named ranges"
    CLEAN ||--o{ SUMMARY : "SUMIFS over c* named ranges"
    CLEAN ||--o{ DASHBOARD : "KPI cards use SUMIFS"
    CLEAN ||--o{ MODERN_FUNCTIONS : "UNIQUE, FILTER over cCountry, cType"
    LISTS ||--o{ DASHBOARD : "lst* named ranges feed dropdowns"
    LISTS ||--o{ SUMMARY : "lstCountry feeds the lookup"
    DASHBOARD ||--o{ CALC : "Sel* cells feed Crit* cells"
    CALC ||--o{ DASHBOARD : "charts read Calc tables"
    CALC ||--o{ MODERN_FUNCTIONS : "XLOOKUP into country table"
```
