"""Stage 2: open stage1.xlsx in real Excel, add dynamic-array formulas,
recalculate, scan for errors, export dashboard PNGs, save the final file.

Usage: python scripts/finish_in_excel.py _work/stage1.xlsx excel-sales-dashboard.xlsx screenshots
Needs Windows, desktop Excel 365 and pywin32.
"""
import os
import sys

import win32com.client as win32

ERR_VALUES = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?",
              -2146826288: "#NULL!", -2146826252: "#NUM!", -2146826265: "#REF!",
              -2146826273: "#VALUE!"}


def scan_errors(wb):
    found = {}
    for ws in wb.Worksheets:
        used = ws.UsedRange
        vals = used.Value2
        if vals is None or not isinstance(vals, tuple):
            continue
        for i, row in enumerate(vals, start=1):
            for j, v in enumerate(row, start=1):
                if isinstance(v, int) and v in ERR_VALUES:
                    cell = used.Cells(i, j).Address(False, False)
                    found.setdefault((ws.Name, ERR_VALUES[v]), []).append(cell)
    return found


def export_dashboard(excel, ws, path):
    """Dashboard -> one-page PDF (print area is set) -> PNG. CopyPicture is blank when Excel is hidden."""
    import fitz
    pdf = path[:-4] + ".pdf"
    ws.ExportAsFixedFormat(0, pdf)
    doc = fitz.open(pdf)
    doc[0].get_pixmap(dpi=150).save(path)
    doc.close()
    os.remove(pdf)


def main(src, dst, shots):
    src, dst, shots = map(os.path.abspath, (src, dst, shots))
    os.makedirs(shots, exist_ok=True)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(src)

        # Modern Functions sheet: the dynamic-array formulas, written by real Excel.
        mf = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        mf.Name = "Modern Functions"
        mf.Move(None, wb.Worksheets(wb.Worksheets.Count))  # After= kwarg is ignored through COM
        mf.Range("A1").Value = "Dynamic-array formulas (Excel 365 / 2021+)"
        mf.Range("A2").Value = ("These spill. If your Excel is older, they show #NAME?. "
                                "The rest of the workbook does not depend on this sheet.")
        mf.Range("A4").Value = "Pick a country"
        mf.Range("B4").Value = "Germany"
        mf.Range("B4").Validation.Add(3, 1, 1, "=lstCountry")
        mf.Range("A5").Value = "Net revenue (XLOOKUP into the country table)"
        mf.Range("B5").Formula2 = "=XLOOKUP(B4,Calc!O9:O1000,Calc!P9:P1000,\"not found\")"
        mf.Range("A7").Value = "Distinct countries, sorted A to Z (UNIQUE + SORT)"
        mf.Range("A8").Formula2 = "=SORT(UNIQUE(cCountry))"
        mf.Range("C7").Value = "Sale lines for the picked country, top 15 by revenue (FILTER + SORTBY + TAKE)"
        mf.Range("C8").Formula2 = ('=TAKE(SORTBY(FILTER(CHOOSE({1,2,3,4},cInvoice,cDesc,cQty,cRevenue),'
                                   '(cCountry=B4)*(cType="Sale")),'
                                   'FILTER(cRevenue,(cCountry=B4)*(cType="Sale")),-1),15)')
        mf.Range("C8").Select()
        mf.Columns("A:A").ColumnWidth = 48
        mf.Columns("B:B").ColumnWidth = 16
        mf.Columns("C:F").ColumnWidth = 22
        mf.Range("A1").Font.Bold = True

        excel.CalculateFull()
        errs = scan_errors(wb)

        dash = wb.Worksheets("Dashboard")
        export_dashboard(excel, dash, os.path.join(shots, "dashboard-all.png"))
        # change dropdowns like a user would, then re-export
        dash.Range("B5").Value = "Germany"
        dash.Range("D5").Value = "All"
        excel.Calculate()
        export_dashboard(excel, dash, os.path.join(shots, "dashboard-germany.png"))
        germany = {a: dash.Range(a).Value2 for a in ("B8", "D8", "F8", "H8", "J8", "L8")}
        dash.Range("B5").Value = "All"
        excel.Calculate()
        allv = {a: dash.Range(a).Value2 for a in ("B8", "D8", "F8", "H8", "J8", "L8")}

        errs_after = scan_errors(wb)
        print("KPIs All    :", allv)
        print("KPIs Germany:", germany)
        print("Modern B5   :", mf.Range("B5").Value2)
        print("Spill A8    :", mf.Range("A8").SpillingToRange.Address if mf.Range("A8").HasSpill else "no spill")
        print("Spill C8    :", mf.Range("C8").SpillingToRange.Address if mf.Range("C8").HasSpill else "no spill")
        print("Errors      :", {k: v[:5] for k, v in errs.items()} or "none")
        print("Errors after:", {k: v[:5] for k, v in errs_after.items()} or "none")
        dash.Activate()
        dash.Range("B5").Select()
        wb.SaveAs(dst, 51)
        wb.Close(False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main(*sys.argv[1:4])
