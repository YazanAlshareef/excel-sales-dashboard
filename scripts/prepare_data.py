"""Step 1: take a reproducible sample of whole invoices from the UCI Online Retail file.

Usage: python scripts/prepare_data.py "path/to/Online Retail.xlsx"
Output: data/raw_sample.csv  (rows copied as-is, nothing cleaned)
"""
import sys
import pandas as pd

SEED = 42
N_INVOICES = 400  # whole invoices, so duplicate lines and returns stay realistic

src = sys.argv[1]
df = pd.read_excel(src, dtype={"InvoiceNo": str, "StockCode": str})
invoices = df["InvoiceNo"].drop_duplicates().sample(N_INVOICES, random_state=SEED)
sample = df[df["InvoiceNo"].isin(invoices)]  # keeps original row order
sample.to_csv("data/raw_sample.csv", index=False)
print(f"source rows {len(df)}, sample rows {len(sample)}, invoices {N_INVOICES}")
