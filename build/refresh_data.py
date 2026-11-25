#!/usr/bin/env python3
"""Refresh index data and rebuild the offline simulator.

Fetches the latest official Price + TRI history for all indices from
niftyindices.com, saves a dated JSON snapshot (so you keep history), updates
data_snapshots/latest.json, and regenerates the single-file offline app.

Usage:
  python build/refresh_data.py            # fetch fresh data + rebuild HTML
  python build/refresh_data.py --from-cache   # rebuild HTML from latest snapshot (offline, no fetch)
"""
import json, os, sys
from datetime import datetime, date, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SNAP_DIR = os.path.join(HERE, "data_snapshots")
TEMPLATE = os.path.join(HERE, "template.html")
# write index.html at the repo root (GitHub Pages); also update the preview copy if present
OUTPUTS = [os.path.join(ROOT, "index.html")]
_preview = os.path.join(ROOT, "frontend", "public", "simulator.html")
if os.path.isdir(os.path.dirname(_preview)):
    OUTPUTS.append(_preview)

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
BASE = "https://www.niftyindices.com"

# key -> (label, official long name)
COLS = [("n50", "Nifty 50", "Nifty 50"),
        ("n500", "Nifty 500", "Nifty 500"),
        ("mom30", "Nifty 200 Momentum 30", "NIFTY200 MOMENTUM 30"),
        ("mom50", "Nifty 500 Momentum 50", "Nifty500 Momentum 50"),
        ("mmq50", "Multicap Mom Quality 50", "Nifty500 Multicap Momentum Quality 50"),
        ("q50", "Nifty 500 Quality 50", "Nifty500 Quality 50"),
        ("v50", "Nifty 500 Value 50", "Nifty500 Value 50"),
        ("v20", "Nifty 50 Value 20", "Nifty50 Value 20")]


def fetch_all():
    import requests
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    s.get(BASE + "/reports/historical-data", timeout=30)
    mp = json.loads(s.get("https://liveindexsa.niftyindices.com/assets/json/IndexMapping.json",
                          timeout=30).content.decode("utf-8-sig"))

    def resolve(longn):
        for e in mp:
            if e.get("Index_long_name", "").strip().lower() == longn.lower():
                return e["Index_long_name"], e["Trading_Index_Name"]
        for e in mp:
            if longn.lower() in e.get("Index_long_name", "").lower():
                return e["Index_long_name"], e["Trading_Index_Name"]
        raise SystemExit("index not found in mapping: " + longn)

    def fetch(url, name, idx, dk, vk):
        end = datetime.now().strftime("%d-%b-%Y")
        cinfo = ("{'name':'%s','startDate':'01-Jan-1990','endDate':'%s','indexName':'%s'}"
                 % (name, end, idx))
        r = s.post(BASE + url, data=json.dumps({"cinfo": cinfo}),
                   headers={"Content-Type": "application/json; charset=utf-8",
                            "Referer": BASE + "/reports/historical-data",
                            "X-Requested-With": "XMLHttpRequest", "Origin": BASE}, timeout=120)
        d = json.loads(r.content.decode("utf-8-sig"))
        out = {}
        for x in d:
            v = x.get(vk)
            if v in ("-", "", None):
                continue
            out[datetime.strptime(x[dk].strip(), "%d %b %Y").date()] = round(float(str(v).replace(",", "")), 2)
        return out

    price, tri = {}, {}
    for k, _lbl, longn in COLS:
        name, idx = resolve(longn)
        price[k] = fetch("/BackPage/getHistoricaldatatabletoString", name, idx, "HistoricalDate", "CLOSE")
        tri[k] = fetch("/BackPage/getTotalReturnIndexString", name, idx, "Date", "TotalReturnsIndex")
        print(f"{k}: price={len(price[k])} tri={len(tri[k])}")

    keys = [k for k, _, _ in COLS]
    aset = {k: (set(price[k]) & set(tri[k])) for k in keys}
    starts = {k: min(aset[k]) for k in keys}
    # median start date: indices starting on/before it define the master calendar; later ones are null-filled
    T = sorted(starts.values())[len(keys) // 2]
    long_keys = [k for k in keys if starts[k] <= T]
    short_keys = [k for k in keys if starts[k] > T]
    master = None
    for k in long_keys:
        master = set(aset[k]) if master is None else (master & aset[k])
    master = sorted(master)
    print("master axis:", len(master), master[0], "->", master[-1], "| long:", long_keys, "| short:", short_keys)

    dates = [d.isoformat() for d in master]
    return {"dates": dates,
            "labels": {k: lbl for k, lbl, _ in COLS},
            "order": keys,
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "price": {k: [price[k].get(d) for d in master] for k in keys},
            "tri": {k: [tri[k].get(d) for d in master] for k in keys}}


def save_snapshot(data):
    os.makedirs(SNAP_DIR, exist_ok=True)
    dated = os.path.join(SNAP_DIR, date.today().isoformat() + ".json")
    latest = os.path.join(SNAP_DIR, "latest.json")
    blob = json.dumps(data, separators=(",", ":"))
    for p in (dated, latest):
        with open(p, "w", encoding="utf-8") as f:
            f.write(blob)
    print("snapshot saved:", dated)


def load_latest():
    p = os.path.join(SNAP_DIR, "latest.json")
    if not os.path.exists(p):
        raise SystemExit("no cached snapshot found; run without --from-cache first.")
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def build_html(data):
    data_json = json.dumps(data, separators=(",", ":"))
    html = open(TEMPLATE, encoding="utf-8").read().replace("/*__DATA__*/", "window.NSE_DATA=" + data_json + ";")
    for out in OUTPUTS:
        with open(out, "w", encoding="utf-8") as f:
            f.write(html)
        print("wrote", out, round(len(html) / 1024), "KB")


def main():
    if "--from-cache" in sys.argv:
        data = load_latest()
        print("rebuilding from cache:", len(data["dates"]), "dates,", len(data["order"]), "indices")
    else:
        data = fetch_all()
        save_snapshot(data)
    build_html(data)
    print("done. app:", OUTPUTS[0])


if __name__ == "__main__":
    main()
