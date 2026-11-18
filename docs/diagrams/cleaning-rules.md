# Cleaning rules

Order in which `clean()` applies R01 to R09 to the raw frame.

Sources: [cleaning.py](../../scripts/cleaning.py), [test_cleaning.py](../../tests/test_cleaning.py)

```mermaid
flowchart TD
    A["raw frame from load_raw"] --> R01["R01 drop duplicates over COLS8, keep first"]
    R01 --> R02["R02 drop StockCode in SERVICE_CODES"]
    R02 --> R03["R03 drop Quantity <= 0 when InvoiceNo does not start with C"]
    R03 --> R04["R04 drop UnitPrice <= 0"]
    R04 --> R05["R05 trim Description spaces"]
    R05 --> R06["R06 most common Description per StockCode"]
    R06 --> R07["R07 rename Country using COUNTRY_FIX"]
    R07 --> R08["R08 blank CustomerID to GUEST, others to text"]
    R08 --> R09["R09 blank Description to UNKNOWN"]
    R09 --> OUT["clean frame and log rows sorted by raw row, rule"]
```

Every rule writes its changes to the log through the `add_log` helper. R01 to R04 log one `row removed` entry per dropped row.
