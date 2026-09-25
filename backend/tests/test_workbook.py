"""Backend-style verification tests for Factor_Investment_Simulator_Rolling_Pro.xlsm

Covers iteration 2: real NSE data, dashboard charts, Charts_Data sheet,
Web Refresh button, modCharts/modWebRefresh VBA modules, and NSE endpoint.
"""
import os
import subprocess
import zipfile
import pytest
import openpyxl
import requests

WB_PATH = "/app/Factor_Investment_Simulator_Rolling_Pro.xlsm"
BASE_URL = "https://momentum-value-sim.preview.emergentagent.com"


# ---------- fixtures ----------
@pytest.fixture(scope="module")
def wb():
    return openpyxl.load_workbook(WB_PATH, data_only=True, keep_vba=True)


@pytest.fixture(scope="module")
def series(wb):
    ws = wb["Database_Daily"]
    dates, mom, val, ben = [], [], [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        dates.append(row[0]); mom.append(row[1]); val.append(row[2]); ben.append(row[3])
    return dates, mom, val, ben


@pytest.fixture(scope="module")
def zf():
    return zipfile.ZipFile(WB_PATH)


def _cagr(a, b, yrs):
    return (b / a) ** (1 / yrs) - 1


# ---------- REAL DATA ----------
def test_row_count_real(series):
    dates, *_ = series
    assert 5200 <= len(dates) <= 5400, f"rows={len(dates)}"


def test_date_span(series):
    dates, *_ = series
    assert dates[0].year == 2005 and dates[0].month == 4, dates[0]
    assert dates[-1].year >= 2025, dates[-1]


def test_dates_ascending(series):
    dates, *_ = series
    for i in range(1, len(dates)):
        assert dates[i] > dates[i - 1], f"row {i}"


def test_no_blank_or_nonnumeric(series):
    _, mom, val, ben = series
    for i, (m, v, b) in enumerate(zip(mom, val, ben)):
        assert isinstance(m, (int, float)) and m > 0, f"row {i} mom={m}"
        assert isinstance(v, (int, float)) and v > 0, f"row {i} val={v}"
        assert isinstance(b, (int, float)) and b > 0, f"row {i} ben={b}"


def test_momentum_is_top_performer(series):
    dates, mom, val, ben = series
    yrs = (dates[-1] - dates[0]).days / 365.25
    cm, cv, cb = _cagr(mom[0], mom[-1], yrs), _cagr(val[0], val[-1], yrs), _cagr(ben[0], ben[-1], yrs)
    print(f"CAGR mom={cm:.4f} val={cv:.4f} ben={cb:.4f} yrs={yrs:.2f}")
    assert cm > cv and cm > cb, f"mom {cm:.4f} not top vs val {cv:.4f} ben {cb:.4f}"
    for name, c in [("mom", cm), ("val", cv), ("ben", cb)]:
        assert 0.10 <= c <= 0.25, f"{name} CAGR {c:.4f} outside 10-25%"


# ---------- ENGINE ----------
def test_validate_engine_output():
    res = subprocess.run(
        ["python", "/app/build/validate_engine.py"],
        capture_output=True, text=True, timeout=90,
    )
    out = res.stdout + res.stderr
    assert res.returncode == 0, out
    assert "INSUFFICIENT DATA" in out
    # find first (default) scenario block SIP 5Y 60/40 2012-2024 and check mom avg is highest
    idx = out.find("SIP  5Y 60/40  2012-2024")
    if idx < 0:
        idx = out.find("SIP 5Y 60/40  2012-2024")
    assert idx >= 0, out
    block = out[idx: idx + 800]
    def _avg(tag):
        line = [l for l in block.splitlines() if l.strip().startswith(tag)][0]
        return float(line.split("avg=")[1].split("%")[0].strip())
    a_mom, a_val, a_ben = _avg("mom"), _avg("val"), _avg("ben")
    print(f"engine default avg mom={a_mom} val={a_val} ben={a_ben}")
    assert a_mom > a_val and a_mom > a_ben


# ---------- CHARTS WIRING ----------
def test_chart_files_exist(zf):
    names = zf.namelist()
    for req in ["xl/charts/chart1.xml", "xl/charts/chart2.xml",
                "xl/drawings/drawing1.xml", "xl/drawings/vmlDrawing1.vml"]:
        assert req in names, req


def test_charts_reference_charts_data(zf):
    c1 = zf.read("xl/charts/chart1.xml").decode()
    c2 = zf.read("xl/charts/chart2.xml").decode()
    assert "Charts_Data" in c1
    assert "Charts_Data" in c2


def test_dashboard_has_drawing_and_legacy(zf):
    sh = zf.read("xl/worksheets/sheet1.xml").decode()
    assert "<drawing " in sh, "missing <drawing>"
    assert "<legacyDrawing " in sh, "missing <legacyDrawing>"


# ---------- BUTTONS + DEFINED NAMES ----------
def test_vml_has_four_buttons(zf):
    vml = zf.read("xl/drawings/vmlDrawing1.vml").decode()
    for macro in ["RefreshData", "RunSimulation", "ResetInputs", "WebRefresh"]:
        assert macro in vml, macro


def test_defined_names(zf):
    wbxml = zf.read("xl/workbook.xml").decode()
    for n in ["cfg_WebEnable", "cfg_WebStart", "cfg_WebEnd",
              "cfg_MomFile", "cfg_ValFile", "cfg_BenFile"]:
        assert f'name="{n}"' in wbxml, n
    # at least one in_* input name still present
    assert 'name="in_Amount"' in wbxml


# ---------- CHARTS_DATA SHEET ----------
def test_charts_data_sheet(wb):
    assert "Charts_Data" in wb.sheetnames
    s = wb["Charts_Data"]
    assert s.sheet_state == "hidden", s.sheet_state
    eq_headers = [s.cell(1, c).value for c in range(1, 6)]
    assert eq_headers == ["Date", "Momentum", "Value", "Benchmark", "Blended"], eq_headers
    dist_headers = [s.cell(1, c).value for c in range(8, 13)]
    assert dist_headers == ["Return Band", "Momentum", "Value", "Benchmark", "Blended"], dist_headers
    labels = [s.cell(r, 8).value for r in range(2, 14)]
    assert len(labels) == 12
    assert labels[0] == "< -10%"
    assert labels[-1] == "> 40%"


# ---------- VBA MODULES ----------
def test_vba_modules_and_validate():
    try:
        from pyopenvba import ExcelFile
    except ImportError:
        pytest.skip("pyopenvba not installed")
    xf = ExcelFile(WB_PATH)
    names = set(xf.module_names())
    print(f"modules: {names}")
    expected = {"ThisWorkbook", "modValidation", "modBlendEngine", "modRollingReturns",
                "modDataImport", "modDashboard", "modCharts", "modWebRefresh"}
    missing = expected - names
    assert not missing, f"missing: {missing}"
    errors = xf.validate()
    assert errors == [], errors


# ---------- DOWNLOAD LINKS ----------
def test_download_xlsm():
    r = requests.get(f"{BASE_URL}/Factor_Investment_Simulator_Rolling_Pro.xlsm", timeout=30)
    assert r.status_code == 200
    assert len(r.content) > 150_000, f"size={len(r.content)}"


def test_download_samples_zip():
    r = requests.get(f"{BASE_URL}/sample_nse_exports.zip", timeout=30)
    assert r.status_code == 200
    assert len(r.content) > 0


# ---------- NSE WEB REFRESH ENDPOINT ----------
def test_nse_tri_endpoint_reachable():
    ua = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
          "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
    s = requests.Session()
    s.headers.update({"User-Agent": ua})
    try:
        g = s.get("https://www.niftyindices.com/reports/historical-data", timeout=20)
    except Exception as e:
        pytest.skip(f"NSE unreachable from test network: {e}")
    if g.status_code != 200:
        pytest.skip(f"NSE landing not reachable: {g.status_code}")
    body = {"cinfo": "{'name':'NIFTY200 MOMENTUM 30','startDate':'01-Jan-2024',"
                    "'endDate':'31-Jan-2024','indexName':'Nifty200Momentm30'}"}
    r = s.post(
        "https://www.niftyindices.com/BackPage/getTotalReturnIndexString",
        json=body,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Referer": "https://www.niftyindices.com/reports/historical-data",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://www.niftyindices.com",
        },
        timeout=30,
    )
    assert r.status_code == 200, r.status_code
    assert "TotalReturnsIndex" in r.text, r.text[:200]
