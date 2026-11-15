# Components

Scripts, data files and workbook sheets, and how they connect.

Sources: [prepare_data.py](../../scripts/prepare_data.py), [cleaning.py](../../scripts/cleaning.py), [build_workbook.py](../../scripts/build_workbook.py), [finish_in_excel.py](../../scripts/finish_in_excel.py), [test_cleaning.py](../../tests/test_cleaning.py), [ci.yml](../../.github/workflows/ci.yml)

```mermaid
flowchart LR
    uci[/"Online Retail.xlsx<br/>UCI source, not in repo"/]
    csv[("data/raw_sample.csv")]
    stage1[/"stage1.xlsx"/]
    final[("excel-sales-dashboard.xlsx")]
    shots[/"screenshots/*.png"/]

    subgraph scripts["scripts/"]
        prep["prepare_data.py"]
        build["build_workbook.py"]
        clean["cleaning.py<br/>load_raw, clean, RULES"]
        finish["finish_in_excel.py<br/>win32com + fitz"]
    end

    subgraph tests["tests/ and CI"]
        t["test_cleaning.py"]
        ci[".github/workflows/ci.yml"]
    end

    uci --> prep --> csv
    csv --> build
    build -->|imports| clean
    build --> stage1
    stage1 --> finish
    finish --> final
    finish --> shots
    t -->|imports| clean
    t -->|runs as subprocess| build
    ci -->|pytest| t
    ci -->|runs| build
```
