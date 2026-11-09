"""Build the interactive flood early warning map from the spreadsheet in data/.

Usage:  python build/build.py [data_dir] [output_dir]

Reads the one .xlsx file in data_dir, checks every row, and writes output_dir/index.html:
a single self-contained page that works online (GitHub Pages) and offline.
If any row has a problem, it prints what to fix and exits with an error, so the live map
is left unchanged.
"""
import json, re, sys, glob, os
from datetime import date
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "data")
OUT_DIR = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "..", "site")

# Uganda with a small margin; anything outside is almost certainly a typo or swapped lat/lon.
LAT_RANGE, LON_RANGE = (-1.6, 4.4), (29.4, 35.1)

errors = []
def fail(msg): errors.append(msg)

# ---------- find the spreadsheet ----------
files = sorted(f for f in glob.glob(os.path.join(DATA_DIR, "*.xlsx")) if not os.path.basename(f).startswith("~$"))
if len(files) != 1:
    print(f"ERROR: expected exactly one .xlsx file in {DATA_DIR}, found {len(files)}: "
          f"{', '.join(os.path.basename(f) for f in files) or 'none'}.\n"
          "Keep only the current spreadsheet in the data folder.")
    sys.exit(1)
XLSX = files[0]
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["Monitoring Sites"] if "Monitoring Sites" in wb.sheetnames else wb.worksheets[0]

# ---------- columns, matched by header text so column order can change ----------
header = [str(c.value or "").strip().lower() for c in ws[1]]
def col(*keys, required=True):
    for i, h in enumerate(header):
        if any(k in h for k in keys): return i
    if required: fail(f"Missing column: a header containing '{keys[0]}' was not found in row 1.")
    return None
C = dict(id=col("site id"), district=col("district"), place=col("location", "river reach"),
         tech=col("sensor", "technology", "type"), dms=col("gps coordinates", required=False),
         lat=col("latitude"), lon=col("longitude"), short=col("short name", required=False))
if errors:
    print("ERROR in spreadsheet layout:\n- " + "\n- ".join(errors)); sys.exit(1)

def kind(t):
    l = t.lower()
    if "radio" in l: return "radio"
    if "graduated" in l or "staff" in l or "river level gauge" in l: return "gauge"
    if any(k in l for k in ("ott", "radar", "telemetr", "telestrial", "automatic", "sensor", "station")): return "station"
    return None

def tech(t):
    l = t.strip().lower()
    if "ott" in l: return "OTT RLS 500 radar water level sensor"
    if "radar" in l: return "Radar water level sensor"
    if "telestrial" in l or "telemetr" in l: return "Automatic telemetric water level sensor"
    if "graduated" in l or "staff" in l: return "Graduated river level gauge (staff gauge)"
    if "radio" in l: return "Community radio"
    return t.strip()

def clean(desc, district):
    d = desc.strip()
    d = re.sub(r"\[text clipped in atlas\]", "", d)
    d = re.sub(r"\((approximate[^)]*)\)", "", d)
    d = re.sub(r"^%s District,?\s*" % re.escape(district), "", d, flags=re.I)
    d = d.replace("(delta, near Lake A ", "(delta, near Lake Albert)").replace("(delta, near Lake A", "(delta, near Lake Albert)")
    d = re.sub(r"^R\. ", "River ", d)
    d = re.sub(r"\s+\)", ")", d); d = re.sub(r"\(\s+", "(", d)
    d = re.sub(r"\s{2,}", " ", d).strip(" ,")
    if d.count("(") > d.count(")"): d += ")"
    return d.replace("))", ")")

def auto_short(place):
    s = re.sub(r"^(River|R\.)\s+", "", place)
    s = re.sub(r"\s+(area|sub-county)$", "", s, flags=re.I)
    m = re.match(r"([^()]+)\(([^)]+)\)", s)
    if m:
        main, extra = m.group(1).strip(), m.group(2).split(",")[0].strip()
        s = main if extra.lower() in ("downstream", "lower reach", "upstream") else f"{main}, {extra}"
    return s.strip()[:28]

def to_dms(v, pos, neg):
    h = pos if v >= 0 else neg; v = abs(v); d = int(v); m = int((v - d) * 60); s = (v - d - m / 60) * 3600
    return f"{d}°{m:02d}'{s:04.1f}\"{h}"

# ---------- rows ----------
sites, seen = [], {}
for rn, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    if all(v is None or str(v).strip() == "" for v in r): continue
    g = lambda k: r[C[k]] if C[k] is not None and C[k] < len(r) else None
    where = f"Row {rn}"
    try: sid = int(g("id"))
    except (TypeError, ValueError): fail(f"{where}: Site ID '{g('id')}' is not a whole number."); continue
    where = f"Row {rn} (Site ID {sid})"
    if sid in seen: fail(f"{where}: Site ID {sid} is used twice (also row {seen[sid]})."); continue
    seen[sid] = rn
    district = str(g("district") or "").strip()
    if not district: fail(f"{where}: District is empty.")
    t_raw = str(g("tech") or "").strip(); typ = kind(t_raw)
    if not typ: fail(f"{where}: Sensor type '{t_raw}' is not recognised. Use a type that contains "
                     "'Community Radio', 'Graduated river level gauge', 'Radar', 'OTT' or 'telemetric'.")
    try: lat, lon = float(g("lat")), float(g("lon"))
    except (TypeError, ValueError):
        fail(f"{where}: Latitude/Longitude must be decimal numbers (got '{g('lat')}', '{g('lon')}')."); continue
    if not (LAT_RANGE[0] <= lat <= LAT_RANGE[1] and LON_RANGE[0] <= lon <= LON_RANGE[1]):
        hint = " Latitude and longitude look swapped." if (LAT_RANGE[0] <= lon <= LAT_RANGE[1] and LON_RANGE[0] <= lat <= LON_RANGE[1]) else ""
        fail(f"{where}: coordinates {lat}, {lon} are outside Uganda.{hint}")
        continue
    place = clean(str(g("place") or ""), district) or f"Site {sid}"
    short = str(g("short") or "").strip() or auto_short(place)
    dms = str(g("dms") or "").strip() or f"{to_dms(lat, 'N', 'S')} {to_dms(lon, 'E', 'W')}"
    if typ:
        sites.append(dict(id=sid, district=district, type=typ, tech=tech(t_raw), place=place, short=short,
                          dms=dms, lat=round(lat, 6), lon=round(lon, 6)))

if not sites and not errors: fail("The spreadsheet has no site rows.")
if errors:
    print(f"ERROR: the map was NOT updated. Fix these in {os.path.basename(XLSX)} and upload it again:\n- " + "\n- ".join(errors))
    sys.exit(1)

# ---------- assemble ----------
base = json.load(open(os.path.join(HERE, "basemap.json")))
names = {f["properties"]["name"].lower(): f for f in base["districts"]["features"]}
focus = {s["district"].lower() for s in sites}
for f in base["districts"]["features"]:
    f["properties"]["focus"] = f["properties"]["name"].lower() in focus
extra = []
for d in sorted({s["district"] for s in sites}):
    if d.lower() not in names:   # district newer than the boundary data: label below its sites
        ds = [s for s in sites if s["district"] == d]
        extra.append([d, round(min(s["lat"] for s in ds) - 0.19, 4), round(sum(s["lon"] for s in ds) / len(ds), 4)])
data = dict(base, sites=sorted(sites, key=lambda s: s["id"]), extraLabels=extra)
blob = json.dumps(data, separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")

read = lambda n: open(os.path.join(HERE, n), encoding="utf-8").read()
page = read("template.html")
page = page.replace("html,body{height:100%}", "html,body{height:100%}\n[hidden]{display:none!important}\nbody{margin:0}")
page = page.replace("/*__LEAFLET_CSS__*/", read("leaflet.css"))
page = page.replace("<!--__ICONS__-->", read("icons.html"))
page = page.replace("<!--__LOGO__-->", read("logo.svg"))
page = page.replace("/*__DATA__*/", blob)
page = page.replace("<!--__LEAFLET_JS__-->", "<script>" + read("leaflet.js") + "</script>")
page = page.replace("Site data: project GPS survey (GPS_DATA_COORDINATES).",
                    f"Site data: {os.path.basename(XLSX)}. Map updated {date.today():%d %B %Y}.")
page = ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover"></head><body>'
        + page + "</body></html>")
os.makedirs(OUT_DIR, exist_ok=True)
open(os.path.join(OUT_DIR, "index.html"), "w", encoding="utf-8").write(page)
open(os.path.join(OUT_DIR, ".nojekyll"), "w").close()

counts = {t: sum(s["type"] == t for s in sites) for t in ("gauge", "station", "radio")}
print(f"OK: built map from {os.path.basename(XLSX)} with {len(sites)} sites in "
      f"{len({s['district'] for s in sites})} districts "
      f"({counts['gauge']} river gauges, {counts['station']} automatic stations, {counts['radio']} community radios).")
