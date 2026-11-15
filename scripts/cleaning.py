"""The nine cleaning rules (R01 to R09), kept apart from the workbook code so they can be tested."""
import pandas as pd

SERVICE_CODES = {"POST", "DOT", "M", "C2", "AMAZONFEE", "BANK CHARGES", "CRUK", "S", "B", "PADS", "D"}
COUNTRY_FIX = {"EIRE": "Ireland", "RSA": "South Africa", "USA": "United States"}
RULES = [
    ("R01", "Remove exact duplicate row", "Row removed. All 8 fields equal to an earlier row; first copy kept."),
    ("R02", "Remove non-product line", "Row removed. StockCode is a service code: " + ", ".join(sorted(SERVICE_CODES)) + "."),
    ("R03", "Remove stock adjustment", "Row removed. Quantity <= 0 on an invoice that does not start with C (not a customer return)."),
    ("R04", "Remove zero or negative price", "Row removed. UnitPrice <= 0."),
    ("R05", "Trim spaces in Description", "Leading/trailing spaces cut, repeated inner spaces collapsed to one."),
    ("R06", "Standardise Description per StockCode", "Description replaced by the most common spelling for that StockCode."),
    ("R07", "Rename country", "EIRE -> Ireland (also RSA, USA if present)."),
    ("R08", "Fill missing CustomerID", "Blank CustomerID set to GUEST. Other IDs stored as text."),
    ("R09", "Fill missing Description", "Blank Description set to UNKNOWN."),
]
REMOVE_RULES = ["R01", "R02", "R03", "R04"]

COLS8 = ["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice", "CustomerID", "Country"]


def load_raw(path):
    """Read the raw sample CSV. Stops with a clear message if the file or a column is missing."""
    try:
        raw = pd.read_csv(path, dtype={"InvoiceNo": str, "StockCode": str})
    except FileNotFoundError:
        raise SystemExit(f"{path} not found. Run scripts/prepare_data.py first.")
    missing = [c for c in COLS8 if c not in raw.columns]
    if missing:
        raise SystemExit(f"{path} is missing columns: {', '.join(missing)}")
    raw["InvoiceDate"] = pd.to_datetime(raw["InvoiceDate"])
    raw["raw_row"] = raw.index + 2  # row number on the Raw sheet
    return raw


def clean(raw):
    """Apply R01 to R09 to a frame from load_raw. Returns (clean frame, log rows).

    A log row is (raw_row, invoice, stock, rule, field, old, new), sorted by raw row then rule.
    """
    log = []

    def add_log(rows, rule, field, old, new):
        for _, r in rows.iterrows():
            log.append((int(r.raw_row), r.InvoiceNo, r.StockCode, rule, field,
                        old(r) if callable(old) else old, new(r) if callable(new) else new))

    df = raw.copy()

    def drop(mask, rule):
        nonlocal df
        hit = df[mask(df)]
        add_log(hit, rule, "row", lambda r: f"Qty {r.Quantity}, Price {r.UnitPrice}", "row removed")
        df = df.drop(hit.index)

    drop(lambda d: d.duplicated(subset=COLS8, keep="first"), "R01")
    drop(lambda d: d.StockCode.isin(SERVICE_CODES), "R02")
    drop(lambda d: (d.Quantity <= 0) & ~d.InvoiceNo.str.startswith("C"), "R03")
    drop(lambda d: d.UnitPrice <= 0, "R04")

    # R05 trim
    desc_old = df.Description.copy()
    desc_new = df.Description.str.replace(r"\s+", " ", regex=True).str.strip()
    chg = df[desc_old.notna() & (desc_old != desc_new)]
    add_log(chg, "R05", "Description", lambda r: repr(desc_old[r.name]), lambda r: repr(desc_new[r.name]))
    df["Description"] = desc_new

    # R06 most common description per stock code (ties: alphabetical)
    mode = (df.dropna(subset=["Description"]).groupby("StockCode")["Description"]
            .agg(lambda s: sorted(s.value_counts().loc[lambda v: v == v.max()].index)[0]))
    std = df.StockCode.map(mode)
    chg = df[df.Description.notna() & std.notna() & (df.Description != std)]
    add_log(chg, "R06", "Description", lambda r: r.Description, lambda r: std[r.name])
    df.loc[chg.index, "Description"] = std[chg.index]

    # R07 country
    chg = df[df.Country.isin(COUNTRY_FIX)]
    add_log(chg, "R07", "Country", lambda r: r.Country, lambda r: COUNTRY_FIX[r.Country])
    df["Country"] = df.Country.replace(COUNTRY_FIX)

    # R08 customer id -> text or GUEST
    chg = df[df.CustomerID.isna()]
    add_log(chg, "R08", "CustomerID", "(blank)", "GUEST")
    df["CustomerID"] = df.CustomerID.map(lambda v: "GUEST" if pd.isna(v) else str(int(v)))

    # R09 description blank
    chg = df[df.Description.isna()]
    add_log(chg, "R09", "Description", "(blank)", "UNKNOWN")
    df["Description"] = df.Description.fillna("UNKNOWN")

    df = df.reset_index(drop=True)
    log.sort(key=lambda t: (t[0], t[3]))
    return df, log
