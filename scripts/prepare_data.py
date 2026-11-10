"""Step 1: take a reproducible sample of whole invoices from the UCI Online Retail file.

Usage: python scripts/prepare_data.py "path/to/Online Retail.xlsx"
Output: data/raw_sample.csv  (rows copied as-is, nothing cleaned)
The 8,495-row sample only comes out with the full 23 MB source file and seed 42.
The source is not stored in this repo; download it from the UCI repository.
"""
import sys
from pathlib import Path
import pandas as pd

SEED = 42
N_INVOICES = 400  # whole invoices, so duplicate lines and returns stay realistic

src = sys.argv[1]
df = pd.read_excel(src, dtype={"InvoiceNo": str, "StockCode": str})
invoices = df["InvoiceNo"].drop_duplicates().sample(N_INVOICES, random_state=SEED)
sample = df[df["InvoiceNo"].isin(invoices)]  # keeps original row order
out = Path(__file__).resolve().parent.parent / "data" / "raw_sample.csv"
sample.to_csv(out, index=False)
print(f"source rows {len(df)}, sample rows {len(sample)}, invoices {N_INVOICES}")
