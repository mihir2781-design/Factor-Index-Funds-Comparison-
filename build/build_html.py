#!/usr/bin/env python3
"""Build a single self-contained offline HTML factor-investment simulator with
official NSE data (Price + TRI) for 6 indices embedded. No CDN, no backend."""
import json, requests
from datetime import datetime, timezone

UA=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
BASE="https://www.niftyindices.com"
s=requests.Session(); s.headers.update({"User-Agent":UA})
s.get(BASE+"/reports/historical-data", timeout=30)
mp=json.loads(s.get("https://liveindexsa.niftyindices.com/assets/json/IndexMapping.json",timeout=30).content.decode("utf-8-sig"))

def resolve(longn):
    for e in mp:
        if e.get("Index_long_name","").strip().lower()==longn.lower(): return e["Index_long_name"],e["Trading_Index_Name"]
    for e in mp:
        if longn.lower() in e.get("Index_long_name","").lower(): return e["Index_long_name"],e["Trading_Index_Name"]
    raise SystemExit("not found "+longn)

# key -> (label, official long name)
COLS=[("n50","Nifty 50","Nifty 50"),
      ("n500","Nifty 500","Nifty 500"),
      ("mom30","Nifty 200 Momentum 30","NIFTY200 MOMENTUM 30"),
      ("mom50","Nifty 500 Momentum 50","Nifty500 Momentum 50"),
      ("mmq50","Multicap Mom Quality 50","Nifty500 Multicap Momentum Quality 50"),
      ("q50","Nifty 500 Quality 50","Nifty500 Quality 50"),
      ("v50","Nifty 500 Value 50","Nifty500 Value 50"),
      ("v20","Nifty 50 Value 20","Nifty50 Value 20")]

def fetch(url,name,idx,dk,vk):
    cinfo="{'name':'%s','startDate':'01-Jan-1990','endDate':'30-Sep-2026','indexName':'%s'}"%(name,idx)
    r=s.post(BASE+url,data=json.dumps({"cinfo":cinfo}),
      headers={"Content-Type":"application/json; charset=utf-8","Referer":BASE+"/reports/historical-data","X-Requested-With":"XMLHttpRequest","Origin":BASE},timeout=120)
    d=json.loads(r.content.decode("utf-8-sig")); out={}
    for x in d:
        v=x.get(vk)
        if v in ("-","",None): continue
        out[datetime.strptime(x[dk].strip(),"%d %b %Y").date()]=round(float(str(v).replace(",","")),2)
    return out

price={}; tri={}
for k,_lbl,longn in COLS:
    name,idx=resolve(longn)
    price[k]=fetch("/BackPage/getHistoricaldatatabletoString",name,idx,"HistoricalDate","CLOSE")
    tri[k]  =fetch("/BackPage/getTotalReturnIndexString",name,idx,"Date","TotalReturnsIndex")
    print(f"{k}: price={len(price[k])} tri={len(tri[k])}")

# master axis = intersection of long-history indices; late-starting indices null-filled
keys=[k for k,_,_ in COLS]
aset={k:(set(price[k])&set(tri[k])) for k in keys}
starts={k:min(aset[k]) for k in keys}
T=sorted(starts.values())[len(keys)//2]
long_keys=[k for k in keys if starts[k]<=T]
short_keys=[k for k in keys if starts[k]>T]
common=None
for k in long_keys:
    common=set(aset[k]) if common is None else (common&aset[k])
common=sorted(common)
print("master axis:",len(common),common[0],"->",common[-1],"| long:",long_keys,"| short:",short_keys)

dates=[d.isoformat() for d in common]
DATA={"dates":dates,
      "labels":{k:lbl for k,lbl,_ in COLS},
      "order":keys,
      "generated_at":datetime.now(timezone.utc).isoformat(),
      "price":{k:[price[k].get(d) for d in common] for k in keys},
      "tri":{k:[tri[k].get(d) for d in common] for k in keys}}
data_json=json.dumps(DATA,separators=(",",":"))

html=open("/app/build/template.html",encoding="utf-8").read().replace("/*__DATA__*/","window.NSE_DATA="+data_json+";")
for out in ("/app/factor_simulator.html","/app/frontend/public/simulator.html"):
    with open(out,"w",encoding="utf-8") as f: f.write(html)
    print("wrote",out, round(len(html)/1024),"KB")
