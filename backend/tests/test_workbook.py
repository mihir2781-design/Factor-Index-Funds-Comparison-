"""Backend-style verification tests for Factor_Investment_Simulator_Rolling_Pro.xlsm"""
import os
import subprocess
import pytest
import openpyxl
import requests

WB_PATH = "/app/Factor_Investment_Simulator_Rolling_Pro.xlsm"
BASE_URL = "https://momentum-value-sim.preview.emergentagent.com"


@pytest.fixture(scope="module")
def ws():
    wb = openpyxl.load_workbook(WB_PATH, data_only=True, keep_vba=True)
    return wb["Database_Daily"]


@pytest.fixture(scope="module")
def series(ws):
    dates, mom, val, ben = [], [], [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        dates.append(row[0])
        mom.append(row[1])
        val.append(row[2])
        ben.append(row[3])
    return dates, mom, val, ben


def cagr(first, last, years):
    return (last / first) ** (1 / years) - 1


# --- BUG FIX: Momentum must be top performer ---
def test_momentum_is_top_performer(series):
    dates, mom, val, ben = series
    years = (dates[-1] - dates[0]).days / 365.25
    c_mom = cagr(mom[0], mom[-1], years)
    c_val = cagr(val[0], val[-1], years)
    c_ben = cagr(ben[0], ben[-1], years)
    print(f"CAGR Momentum={c_mom:.4f} Value={c_val:.4f} Benchmark={c_ben:.4f} years={years:.2f}")
    assert c_mom > c_val, f"Momentum {c_mom:.4f} not > Value {c_val:.4f}"
    assert c_mom > c_ben, f"Momentum {c_mom:.4f} not > Benchmark {c_ben:.4f}"
    for name, c in [("mom", c_mom), ("val", c_val), ("ben", c_ben)]:
        assert 0.10 <= c <= 0.30, f"{name} CAGR {c:.4f} outside 10-30%"


# --- DATA INTEGRITY ---
def test_row_count(series):
    dates, *_ = series
    assert len(dates) >= 4400, f"Only {len(dates)} rows"


def test_dates_ascending(series):
    dates, *_ = series
    for i in range(1, len(dates)):
        assert dates[i] > dates[i - 1], f"Non-ascending at row {i}"


def test_no_blank_or_nonnumeric(series):
    dates, mom, val, ben = series
    for i, (m, v, b) in enumerate(zip(mom, val, ben)):
        assert isinstance(m, (int, float)) and m > 0, f"row {i} mom={m}"
        assert isinstance(v, (int, float)) and v > 0, f"row {i} val={v}"
        assert isinstance(b, (int, float)) and b > 0, f"row {i} ben={b}"


# --- ROLLING METRIC SANITY via engine port ---
def test_validate_engine_output():
    res = subprocess.run(
        ["python", "/app/build/validate_engine.py"],
        capture_output=True, text=True, timeout=60,
    )
    out = res.stdout + res.stderr
    print(out)
    assert res.returncode == 0, f"engine failed: {out}"
    assert "INSUFFICIENT DATA" in out, "expected insufficient-data scenario"


# --- DOWNLOAD LINKS ---
def test_download_xlsm():
    r = requests.get(f"{BASE_URL}/Factor_Investment_Simulator_Rolling_Pro.xlsm", timeout=30)
    assert r.status_code == 200
    assert len(r.content) > 50_000, f"size={len(r.content)}"


def test_download_samples_zip():
    r = requests.get(f"{BASE_URL}/sample_nse_exports.zip", timeout=30)
    assert r.status_code == 200
    assert len(r.content) > 0


# --- VBA INTEGRITY ---
def test_vba_modules_and_validate():
    try:
        from pyopenvba import ExcelFile
    except ImportError:
        pytest.skip("pyopenvba not installed")
    xf = ExcelFile(WB_PATH)
    names = xf.module_names()
    print(f"modules: {names}")
    expected = {"ThisWorkbook", "modValidation", "modBlendEngine",
                "modRollingReturns", "modDataImport", "modDashboard"}
    missing = expected - set(names)
    assert not missing, f"missing modules: {missing}"
    errors = xf.validate()
    assert errors == [], f"validate errors: {errors}"
