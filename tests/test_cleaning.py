"""Tests for scripts/cleaning.py and the build script's argument check. Run: python -m pytest"""
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from cleaning import clean, load_raw  # noqa: E402


def frame(*rows):
    """Rows are (InvoiceNo, StockCode, Description, Quantity, UnitPrice, CustomerID, Country)."""
    df = pd.DataFrame(rows, columns=["InvoiceNo", "StockCode", "Description", "Quantity",
                                     "UnitPrice", "CustomerID", "Country"])
    df["InvoiceDate"] = pd.Timestamp("2011-01-05 10:00")
    df["raw_row"] = df.index + 2
    return df


def test_duplicate_removed_first_copy_kept():
    raw = frame(("1", "A", "Mug", 1, 2.0, 100.0, "France"), ("1", "A", "Mug", 1, 2.0, 100.0, "France"))
    out, log = clean(raw)
    assert len(out) == 1
    assert [t[3] for t in log] == ["R01"] and log[0][0] == 3  # the second row is the one removed


def test_return_kept_adjustment_removed():
    raw = frame(("C2", "A", "Mug", -1, 2.0, 100.0, "France"),
                ("3", "A", "Mug", -5, 2.0, 100.0, "France"))
    out, log = clean(raw)
    assert list(out.InvoiceNo) == ["C2"]
    assert [t[3] for t in log] == ["R03"]


def test_blank_description_gets_unknown():
    out, log = clean(frame(("1", "A", None, 1, 2.0, 100.0, "France")))
    assert out.Description[0] == "UNKNOWN"
    assert [t[3] for t in log] == ["R09"]


def test_eire_renamed_and_customer_id_becomes_text():
    raw = frame(("1", "A", "Mug", 1, 2.0, 12345.0, "EIRE"), ("2", "A", "Mug", 1, 2.0, None, "France"))
    out, _ = clean(raw)
    assert list(out.Country) == ["Ireland", "France"]
    assert list(out.CustomerID) == ["12345", "GUEST"]


def test_description_trimmed_and_standardised():
    raw = frame(("1", "A", " Mug  ", 1, 2.0, 1.0, "France"), ("2", "A", "Mug", 1, 2.0, 1.0, "France"),
                ("3", "A", "MUG", 1, 2.0, 1.0, "France"))
    out, _ = clean(raw)
    assert set(out.Description) == {"Mug"}


def test_missing_file_gives_clear_message(tmp_path):
    with pytest.raises(SystemExit, match="prepare_data.py"):
        load_raw(tmp_path / "nope.csv")


def test_missing_column_gives_clear_message(tmp_path):
    p = tmp_path / "bad.csv"
    p.write_text("InvoiceNo,StockCode\n1,A\n")
    with pytest.raises(SystemExit, match="missing columns"):
        load_raw(p)


def test_build_script_without_argument_prints_usage():
    r = subprocess.run([sys.executable, str(ROOT / "scripts" / "build_workbook.py")],
                       capture_output=True, text=True)
    assert r.returncode != 0 and "Usage" in r.stderr


def test_sample_rule_counts_match_readme():
    out, log = clean(load_raw(ROOT / "data" / "raw_sample.csv"))
    counts = pd.Series([t[3] for t in log]).value_counts().to_dict()
    assert counts == {"R01": 69, "R02": 51, "R03": 27, "R04": 13,
                      "R05": 1994, "R06": 85, "R07": 65, "R08": 2241}
    assert len(out) == 8335 and len(log) == 4545
