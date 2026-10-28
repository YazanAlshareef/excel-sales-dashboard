"""Step 2: clean data/raw_sample.csv in pandas, then write the workbook with openpyxl.

Usage: python scripts/build_workbook.py _work/stage1.xlsx
Step 3 (scripts/finish_in_excel.py) opens that file in real Excel, adds the
Excel-365-only sheet, recalculates, scans for errors and exports screenshots.

Everything the workbook shows is either:
  * static  - written by this script (Raw, Clean data columns A-I, the log, list keys)
  * live    - an Excel formula (Clean columns J-N, Summary, Calc, Dashboard)
"""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = sys.argv[1]

# ---------------------------------------------------------------- rules
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

raw = pd.read_csv("data/raw_sample.csv", dtype={"InvoiceNo": str, "StockCode": str}, parse_dates=["InvoiceDate"])
raw["raw_row"] = raw.index + 2  # row number on the Raw sheet
log = []  # (raw_row, invoice, stock, rule, field, old, new)


def add_log(rows, rule, field, old, new):
    for _, r in rows.iterrows():
        log.append((int(r.raw_row), r.InvoiceNo, r.StockCode, rule, field,
                    old(r) if callable(old) else old, new(r) if callable(new) else new))


df = raw.copy()
cols8 = ["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice", "CustomerID", "Country"]


def drop(mask, rule):
    global df
    hit = df[mask(df)]
    add_log(hit, rule, "row", lambda r: f"Qty {r.Quantity}, Price {r.UnitPrice}", "row removed")
    df = df.drop(hit.index)


drop(lambda d: d.duplicated(subset=cols8, keep="first"), "R01")
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
df["Month"] = df.InvoiceDate.dt.strftime("%Y-%m")
df["Revenue"] = df.Quantity * df.UnitPrice
months = sorted(df.Month.unique())
countries = df.groupby("Country").Revenue.sum().sort_values(ascending=False).index.tolist()
countries_alpha = sorted(countries)
products = sorted(df.StockCode.unique())
log.sort(key=lambda t: (t[0], t[3]))
print("raw", len(raw), "clean", len(df), "log", len(log), "products", len(products), "countries", len(countries))

# ---------------------------------------------------------------- styles
F = "Arial"
f_norm = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
f_title = Font(name=F, size=20, bold=True, color="1F3864")
f_h2 = Font(name=F, size=14, bold=True, color="1F3864")
f_sub = Font(name=F, size=10, italic=True, color="595959")
f_input = Font(name=F, size=11, bold=True, color="0000FF")
f_kpi = Font(name=F, size=18, bold=True, color="1F3864")
f_kpil = Font(name=F, size=9, bold=True, color="595959")
fill_head = PatternFill("solid", fgColor="1F3864")
fill_input = PatternFill("solid", fgColor="FFF2CC")
fill_card = PatternFill("solid", fgColor="EAF0F8")
fill_tot = PatternFill("solid", fgColor="D9E1F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
MONEY = '#,##0.00;(#,##0.00);-'
MONEY0 = '#,##0;(#,##0);-'
INT = '#,##0;(#,##0);-'

wb = Workbook()
wb._fonts = IndexedList([f_norm])  # make Arial the default font for every cell
wb._named_styles["Normal"].font = f_norm

ws_dash = wb.active
ws_dash.title = "Dashboard"
ws_sum = wb.create_sheet("Summary")
ws_clean = wb.create_sheet("Clean")
ws_log = wb.create_sheet("Cleaning Log")
ws_raw = wb.create_sheet("Raw")
ws_calc = wb.create_sheet("Calc")
ws_list = wb.create_sheet("Lists")


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def header(ws, row, col, labels):
    for j, t in enumerate(labels):
        c = ws.cell(row, col + j, t)
        c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box


# ---------------------------------------------------------------- Raw (untouched)
header(ws_raw, 1, 1, cols8)
for i, r in enumerate(raw[cols8].itertuples(index=False), start=2):
    ws_raw.cell(i, 1, r[0])
    ws_raw.cell(i, 2, r[1])
    if not pd.isna(r[2]):
        ws_raw.cell(i, 3, r[2])
    ws_raw.cell(i, 4, int(r[3]))
    ws_raw.cell(i, 5, r[4].to_pydatetime()).number_format = "yyyy-mm-dd hh:mm"
    ws_raw.cell(i, 6, float(r[5]))
    if not pd.isna(r[6]):
        ws_raw.cell(i, 7, float(r[6]))
    ws_raw.cell(i, 8, r[7])
for col, w in zip("ABCDEFGH", [11, 12, 40, 9, 17, 10, 11, 18]):
    ws_raw.column_dimensions[col].width = w
ws_raw.freeze_panes = "A2"
ws_raw.auto_filter.ref = f"A1:H{len(raw) + 1}"

# ---------------------------------------------------------------- Clean
n = len(df)
last = n + 1
clean_heads = ["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice",
               "CustomerID", "Country", "RawRow", "Type", "Revenue", "Month", "CustType", "OrderFlag"]
header(ws_clean, 1, 1, clean_heads)
for i, r in enumerate(df.itertuples(index=False), start=2):
    ws_clean.cell(i, 1, r.InvoiceNo)
    ws_clean.cell(i, 2, r.StockCode)
    ws_clean.cell(i, 3, r.Description)
    ws_clean.cell(i, 4, int(r.Quantity))
    ws_clean.cell(i, 5, r.InvoiceDate.to_pydatetime()).number_format = "yyyy-mm-dd hh:mm"
    ws_clean.cell(i, 6, float(r.UnitPrice)).number_format = "0.00"
    ws_clean.cell(i, 7, r.CustomerID)
    ws_clean.cell(i, 8, r.Country)
    ws_clean.cell(i, 9, int(r.raw_row))
    ws_clean.cell(i, 10, f'=IF(LEFT(A{i},1)="C","Return","Sale")')
    ws_clean.cell(i, 11, f"=D{i}*F{i}").number_format = MONEY
    ws_clean.cell(i, 12, f'=TEXT(E{i},"yyyy-mm")')
    ws_clean.cell(i, 13, f'=IF(G{i}="GUEST","Guest","Registered")')
    ws_clean.cell(i, 14, f'=IF(AND(J{i}="Sale",COUNTIF($A$2:A{i},A{i})=1),1,0)')
for col, w in zip("ABCDEFGHIJKLMN", [11, 12, 40, 9, 17, 10, 11, 18, 8, 9, 12, 9, 12, 10]):
    ws_clean.column_dimensions[col].width = w
ws_clean.freeze_panes = "A2"
tbl = Table(displayName="tblClean", ref=f"A1:N{last}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
ws_clean.add_table(tbl)
for nm, col in [("cInvoice", "A"), ("cStock", "B"), ("cDesc", "C"), ("cQty", "D"), ("cDate", "E"),
                ("cPrice", "F"), ("cCust", "G"), ("cCountry", "H"), ("cType", "J"), ("cRevenue", "K"),
                ("cMonth", "L"), ("cCustType", "M"), ("cOrderFlag", "N")]:
    name(nm, f"Clean!${col}$2:${col}${last}")

# ---------------------------------------------------------------- Lists
header(ws_list, 1, 1, ["Country list", "Month list", "Customer type list", "Metric list"])
for i, v in enumerate(["All"] + countries_alpha, start=2):
    ws_list.cell(i, 1, v)
for i, v in enumerate(["All"] + months, start=2):
    ws_list.cell(i, 2, v)
for i, v in enumerate(["All", "Registered", "Guest"], start=2):
    ws_list.cell(i, 3, v)
for i, v in enumerate(["Revenue", "Units"], start=2):
    ws_list.cell(i, 4, v)
name("lstCountry", f"Lists!$A$2:$A${len(countries_alpha) + 1}")
name("lstMonth", f"Lists!$B$2:$B${len(months) + 1}")
name("lstCust", "Lists!$C$2:$C$4")
name("lstMetric", "Lists!$D$2:$D$3")
for col in "ABCD":
    ws_list.column_dimensions[col].width = 22

# ---------------------------------------------------------------- Dashboard input cells (layout below)
name("SelCountry", "Dashboard!$B$5")
name("SelMonth", "Dashboard!$D$5")
name("SelCust", "Dashboard!$F$5")
name("SelMetric", "Dashboard!$H$5")

# ---------------------------------------------------------------- Calc
ws_calc["A1"] = "Calc - helper tables that feed the Dashboard charts (all live formulas)"
ws_calc["A1"].font = f_bold
for r, (lab, f, nm) in enumerate([("Country criterion", '=IF(SelCountry="All","*",SelCountry)', "CritCountry"),
                                  ("Month criterion", '=IF(SelMonth="All","*",SelMonth)', "CritMonth"),
                                  ("Customer type criterion", '=IF(SelCust="All","*",SelCust)', "CritCust")], start=2):
    ws_calc.cell(r, 1, lab)
    ws_calc.cell(r, 2, f)
    name(nm, f"Calc!$B${r}")
HR = 8  # header row of every block
# monthly block A:D
header(ws_calc, HR, 1, ["Month", "Net revenue", "Orders", "Sales revenue"])
for i, m in enumerate(months, start=HR + 1):
    ws_calc.cell(i, 1, m)
    ws_calc.cell(i, 2, f"=SUMIFS(cRevenue,cMonth,$A{i},cCountry,CritCountry,cCustType,CritCust)").number_format = MONEY0
    ws_calc.cell(i, 3, f"=SUMIFS(cOrderFlag,cMonth,$A{i},cCountry,CritCountry,cCustType,CritCust)").number_format = INT
    ws_calc.cell(i, 4, f'=SUMIFS(cRevenue,cMonth,$A{i},cType,"Sale",cCountry,CritCountry,cCustType,CritCust)').number_format = MONEY0
m_last = HR + len(months)
# full product table S:V
PS, PT, PV, PK = 19, 20, 21, 22
header(ws_calc, HR, PS, ["StockCode", "Description", "Value", "RankKey"])
p0, p1 = HR + 1, HR + len(products)
crit = "cCountry,CritCountry,cMonth,CritMonth,cCustType,CritCust"
for i, p in enumerate(products, start=p0):
    ws_calc.cell(i, PS, p)
    ws_calc.cell(i, PT, f"=INDEX(cDesc,MATCH(${L(PS)}{i},cStock,0))")
    ws_calc.cell(i, PV, f'=IF(SelMetric="Revenue",SUMIFS(cRevenue,cStock,${L(PS)}{i},{crit}),SUMIFS(cQty,cStock,${L(PS)}{i},{crit}))')
    ws_calc.cell(i, PK, f"=${L(PV)}{i}+(100000-ROW())/1E9")
pk_rng = f"${L(PK)}${p0}:${L(PK)}${p1}"
# top 10 products F:J
header(ws_calc, HR, 6, ["Rank", "Pos", "StockCode", "Description", "Value"])
for k in range(1, 11):
    r = HR + k
    ws_calc.cell(r, 6, k)
    ws_calc.cell(r, 7, f"=MATCH(LARGE({pk_rng},F{r}),{pk_rng},0)")
    ws_calc.cell(r, 8, f"=INDEX(${L(PS)}${p0}:${L(PS)}${p1},G{r})")
    ws_calc.cell(r, 9, f'=LEFT(INDEX(${L(PT)}${p0}:${L(PT)}${p1},G{r}),28)')
    ws_calc.cell(r, 10, f"=INDEX(${L(PV)}${p0}:${L(PV)}${p1},G{r})").number_format = MONEY0
# full country table O:Q
CC, CV, CK = 15, 16, 17
header(ws_calc, HR, CC, ["Country", "Value", "RankKey"])
c0, c1 = HR + 1, HR + len(countries)
ccrit = "cMonth,CritMonth,cCustType,CritCust"
for i, c in enumerate(countries, start=c0):
    ws_calc.cell(i, CC, c)
    ws_calc.cell(i, CV, f'=IF(SelMetric="Revenue",SUMIFS(cRevenue,cCountry,${L(CC)}{i},{ccrit}),SUMIFS(cQty,cCountry,${L(CC)}{i},{ccrit}))').number_format = MONEY0
    ws_calc.cell(i, CK, f'=${L(CV)}{i}+(1000-ROW())/1E9-IF(${L(CC)}{i}="United Kingdom",1E12,0)')
ck_rng = f"${L(CK)}${c0}:${L(CK)}${c1}"
# top countries outside the UK, K:N
header(ws_calc, HR, 11, ["Rank", "Pos", "Country", "Value"])
NC = min(8, len(countries) - 1)
for k in range(1, NC + 1):
    r = HR + k
    ws_calc.cell(r, 11, k)
    ws_calc.cell(r, 12, f"=MATCH(LARGE({ck_rng},K{r}),{ck_rng},0)")
    ws_calc.cell(r, 13, f"=INDEX(${L(CC)}${c0}:${L(CC)}${c1},L{r})")
    ws_calc.cell(r, 14, f"=INDEX(${L(CV)}${c0}:${L(CV)}${c1},L{r})").number_format = MONEY0
for c in range(1, 23):
    ws_calc.column_dimensions[L(c)].width = 14
ws_calc.column_dimensions["I"].width = 30
ws_calc.column_dimensions["T"].width = 36
ws_calc.column_dimensions["A"].width = 26

# ---------------------------------------------------------------- Summary (unfiltered, live)
ws = ws_sum
ws["A1"] = "Summary tables (all data, no filters)"
ws["A1"].font = f_h2
ws["A2"] = "Live SUMIFS formulas over the Clean sheet. Row order of the country table was set when the file was built."
ws["A2"].font = f_sub
ws["A4"] = "A. By month"
ws["A4"].font = f_bold
header(ws, 5, 1, ["Month", "Sales revenue", "Returns", "Net revenue", "Orders", "Net units", "Avg order value"])
s0 = 6
for i, m in enumerate(months, start=s0):
    ws.cell(i, 1, m)
    ws.cell(i, 2, f'=SUMIFS(cRevenue,cMonth,$A{i},cType,"Sale")').number_format = MONEY0
    ws.cell(i, 3, f'=SUMIFS(cRevenue,cMonth,$A{i},cType,"Return")').number_format = MONEY0
    ws.cell(i, 4, f"=B{i}+C{i}").number_format = MONEY0
    ws.cell(i, 5, f"=SUMIFS(cOrderFlag,cMonth,$A{i})").number_format = INT
    ws.cell(i, 6, f"=SUMIFS(cQty,cMonth,$A{i})").number_format = INT
    ws.cell(i, 7, f"=IFERROR(B{i}/E{i},0)").number_format = MONEY
s1 = s0 + len(months) - 1
t = s1 + 1
ws.cell(t, 1, "Total")
for col in "BCDEF":
    ws[f"{col}{t}"] = f"=SUM({col}{s0}:{col}{s1})"
    ws[f"{col}{t}"].number_format = MONEY0
ws[f"G{t}"] = f"=IFERROR(B{t}/E{t},0)"
ws[f"G{t}"].number_format = MONEY
for c in range(1, 8):
    ws.cell(t, c).font, ws.cell(t, c).fill = f_bold, fill_tot
ws.cell(t + 1, 1, "Check: monthly net total equals SUM of Clean revenue")
ws.cell(t + 1, 4, f'=IF(ABS(D{t}-SUM(cRevenue))<0.01,"OK","CHECK")').font = f_bold

r0 = t + 4
ws.cell(r0 - 1, 1, "B. By country (sorted by net revenue when built)").font = f_bold
header(ws, r0, 1, ["Country", "Net revenue", "Orders", "Share of net revenue"])
k0 = r0 + 1
k1 = k0 + len(countries) - 1
tc = k1 + 1
for i, c in enumerate(countries, start=k0):
    ws.cell(i, 1, c)
    ws.cell(i, 2, f"=SUMIFS(cRevenue,cCountry,$A{i})").number_format = MONEY0
    ws.cell(i, 3, f"=SUMIFS(cOrderFlag,cCountry,$A{i})").number_format = INT
    ws.cell(i, 4, f"=IFERROR(B{i}/B${tc},0)").number_format = "0.0%"
ws.cell(tc, 1, "Total")
ws[f"B{tc}"] = f"=SUM(B{k0}:B{k1})"
ws[f"B{tc}"].number_format = MONEY0
ws[f"C{tc}"] = f"=SUM(C{k0}:C{k1})"
ws[f"C{tc}"].number_format = INT
ws[f"D{tc}"] = f"=SUM(D{k0}:D{k1})"
ws[f"D{tc}"].number_format = "0.0%"
for c in range(1, 5):
    ws.cell(tc, c).font, ws.cell(tc, c).fill = f_bold, fill_tot

u0 = tc + 3
ws.cell(u0 - 1, 1, "C. By customer type").font = f_bold
header(ws, u0, 1, ["Customer type", "Net revenue", "Orders", "Share of net revenue"])
for i, ct in enumerate(["Registered", "Guest"], start=u0 + 1):
    ws.cell(i, 1, ct)
    ws.cell(i, 2, f"=SUMIFS(cRevenue,cCustType,$A{i})").number_format = MONEY0
    ws.cell(i, 3, f"=SUMIFS(cOrderFlag,cCustType,$A{i})").number_format = INT
    ws.cell(i, 4, f"=IFERROR(B{i}/SUM(cRevenue),0)").number_format = "0.0%"

v0 = u0 + 5
ws.cell(v0 - 1, 1, "D. Country lookup (INDEX/MATCH)").font = f_bold
ws.cell(v0, 1, "Pick a country")
ws.cell(v0, 2, "France").font = f_input
ws.cell(v0, 2).fill = fill_input
ws.cell(v0 + 1, 1, "Net revenue")
ws.cell(v0 + 1, 2, f"=INDEX($B${k0}:$B${k1},MATCH($B${v0},$A${k0}:$A${k1},0))").number_format = MONEY
ws.cell(v0 + 2, 1, "Orders")
ws.cell(v0 + 2, 2, f"=INDEX($C${k0}:$C${k1},MATCH($B${v0},$A${k0}:$A${k1},0))").number_format = INT
dv = DataValidation(type="list", formula1="=lstCountry", allow_blank=False, showErrorMessage=True,
                    errorTitle="Country", error="Pick a country from the list (not All).")
ws.add_data_validation(dv)
dv.add(ws.cell(v0, 2))
name("LookupCountry", f"Summary!$B${v0}")
ws.column_dimensions["A"].width = 30
for col in "BCDEFG":
    ws.column_dimensions[col].width = 18

# ---------------------------------------------------------------- Cleaning Log
ws = ws_log
header(ws, 1, 1, ["LogID", "RawRow", "InvoiceNo", "StockCode", "Rule", "Field", "Old value", "New value / action"])
for i, (rr, inv, sc, rule, field, old, new) in enumerate(log, start=2):
    ws.cell(i, 1, i - 1)
    ws.cell(i, 2, rr)
    ws.cell(i, 3, inv)
    ws.cell(i, 4, sc)
    ws.cell(i, 5, rule)
    ws.cell(i, 6, field)
    ws.cell(i, 7, str(old))
    ws.cell(i, 8, str(new))
lg_last = len(log) + 1
for col, w in zip("ABCDEFGH", [8, 9, 11, 12, 7, 12, 44, 44]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:H{lg_last}"
header(ws, 1, 10, ["Rule", "What", "Effect", "Rows affected"])
for i, (rid, what, eff) in enumerate(RULES, start=2):
    ws.cell(i, 10, rid)
    ws.cell(i, 11, what)
    ws.cell(i, 12, eff)
    ws.cell(i, 13, f"=COUNTIF($E$2:$E${lg_last},J{i})").number_format = INT
rr_ = len(RULES) + 3
rows = [
    ("Rows in Raw", "=COUNTA(Raw!A:A)-1"),
    ("Rows removed (R01-R04)", f"=SUM(M2:M{1 + len(REMOVE_RULES)})"),
    ("Rows in Clean", "=COUNTA(Clean!A:A)-1"),
    ("Reconciles (Raw - removed = Clean)", f'=IF(M{rr_}-M{rr_ + 1}=M{rr_ + 2},"OK","CHECK")'),
]
for i, (lab, f) in enumerate(rows):
    ws.cell(rr_ + i, 11, lab).font = f_bold
    ws.cell(rr_ + i, 13, f).number_format = INT
ws.cell(rr_ + 5, 11, "Return invoices (start with C) are kept and labelled Type = Return, not removed.")
ws.column_dimensions["J"].width = 7
ws.column_dimensions["K"].width = 40
ws.column_dimensions["L"].width = 80
ws.column_dimensions["M"].width = 14

# ---------------------------------------------------------------- Dashboard
ws = ws_dash
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
for c in range(2, 14):
    ws.column_dimensions[L(c)].width = 13
ws["B2"] = "Sales Operations Dashboard"
ws["B2"].font = f_title
ws["B3"] = ("Source: UCI Online Retail dataset (CC BY 4.0), sample of 400 invoices. "
            "Yellow cells are dropdowns.")
ws["B3"].font = f_sub
for col, lab, default, lst in [("B", "Country", "All", "lstCountry"), ("D", "Month", "All", "lstMonth"),
                               ("F", "Customer type", "All", "lstCust"), ("H", "Metric (charts)", "Revenue", "lstMetric")]:
    nxt = L(ord(col) - 64 + 1)
    ws[f"{col}4"] = lab
    ws[f"{col}4"].font = f_kpil
    ws.merge_cells(f"{col}5:{nxt}5")
    ws[f"{col}5"] = default
    for cc in (col, nxt):
        ws[f"{cc}5"].fill, ws[f"{cc}5"].border = fill_input, box
    ws[f"{col}5"].font, ws[f"{col}5"].alignment = f_input, center
    d = DataValidation(type="list", formula1=f"={lst}", allow_blank=False, showErrorMessage=True,
                       errorTitle=lab, error="Pick a value from the dropdown list.")
    ws.add_data_validation(d)
    d.add(ws[f"{col}5"])
ws["J4"] = "Showing"
ws["J4"].font = f_kpil
ws["J5"] = '=SelCountry&" | "&SelMonth&" | "&SelCust&" | "&SelMetric'
ws["J5"].font = f_sub

# KPI cards: value cells are row 8 of B,D,F,H,J,L. D8 = sales revenue, F8 = returns, H8 = orders.
flt = "cCountry,CritCountry,cMonth,CritMonth,cCustType,CritCust"
cards = [
    ("B", "NET REVENUE", f"=SUMIFS(cRevenue,{flt})", MONEY0),
    ("D", "SALES REVENUE", f'=SUMIFS(cRevenue,cType,"Sale",{flt})', MONEY0),
    ("F", "RETURNS VALUE", f'=SUMIFS(cRevenue,cType,"Return",{flt})', MONEY0),
    ("H", "ORDERS", f"=SUMIFS(cOrderFlag,{flt})", INT),
    ("J", "AVG ORDER VALUE", "=IFERROR(D8/H8,0)", MONEY),
    ("L", "RETURN RATE", "=IFERROR(-F8/D8,0)", "0.0%"),
]
for col, lab, f, nf in cards:
    nxt = L(ord(col) - 64 + 1)
    ws.merge_cells(f"{col}7:{nxt}7")
    ws.merge_cells(f"{col}8:{nxt}9")
    ws[f"{col}7"] = lab
    ws[f"{col}7"].font, ws[f"{col}7"].alignment = f_kpil, center
    ws[f"{col}8"] = f
    ws[f"{col}8"].font, ws[f"{col}8"].alignment, ws[f"{col}8"].number_format = f_kpi, center, nf
    for r in (7, 8, 9):
        for cc in (col, nxt):
            ws[f"{cc}{r}"].fill = fill_card

ws["B10"] = "Month charts ignore the Month dropdown; country chart leaves out the United Kingdom (see Summary for its share)."
ws["B10"].font = f_sub


def style_bar(ch, title, w=15.2, h=7.6):
    ch.title = title
    ch.legend = None
    ch.width, ch.height = w, h
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.number_format = "#,##0"
    ch.series[0].graphicalProperties.solidFill = "2F5597"


ch1 = BarChart()
ch1.type = "col"
ch1.add_data(Reference(ws_calc, min_col=2, min_row=HR, max_row=m_last), titles_from_data=True)
ch1.set_categories(Reference(ws_calc, min_col=1, min_row=HR + 1, max_row=m_last))
style_bar(ch1, "Net revenue by month")
ws.add_chart(ch1, "B11")

ch2 = BarChart()
ch2.type = "bar"
ch2.add_data(Reference(ws_calc, min_col=10, min_row=HR, max_row=HR + 10), titles_from_data=True)
ch2.set_categories(Reference(ws_calc, min_col=9, min_row=HR + 1, max_row=HR + 10))
style_bar(ch2, "Top 10 products")
ch2.x_axis.scaling.orientation = "maxMin"
ws.add_chart(ch2, "H11")

ch3 = BarChart()
ch3.type = "bar"
ch3.add_data(Reference(ws_calc, min_col=14, min_row=HR, max_row=HR + NC), titles_from_data=True)
ch3.set_categories(Reference(ws_calc, min_col=13, min_row=HR + 1, max_row=HR + NC))
style_bar(ch3, "Top countries outside the UK")
ch3.x_axis.scaling.orientation = "maxMin"
ws.add_chart(ch3, "B27")

ch4 = LineChart()
ch4.add_data(Reference(ws_calc, min_col=3, min_row=HR, max_row=m_last), titles_from_data=True)
ch4.set_categories(Reference(ws_calc, min_col=1, min_row=HR + 1, max_row=m_last))
ch4.title, ch4.legend, ch4.width, ch4.height = "Orders by month", None, 15.2, 7.6
ch4.x_axis.delete = False
ch4.y_axis.delete = False
ch4.series[0].graphicalProperties.line.solidFill = "C55A11"
ch4.series[0].graphicalProperties.line.width = 28000
ws.add_chart(ch4, "H27")

ws.print_area = "B2:M42"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws.page_margins.left = ws.page_margins.right = 0.4
ws.page_margins.top = ws.page_margins.bottom = 0.4

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
