#!/usr/bin/env python3
"""
metar_extremes.py - monthly METAR extreme-event analysis from Ogimet.

Two independent analyses of the same data:
  A) LOCAL CONDITIONS - tiered criteria (Normal / Moderate / Severe / Very Extreme) for wind,
     rain, visibility, ceiling, dangerous phenomena and heat. Edit the LOCAL dict below.
  B) ICAO ANNEX 3     - SPECI-type criteria: visibility and ceiling bands, onset of weather
     phenomena, wind-change criteria. Edit the ANNEX dict below.

Command line:   python metar_extremes.py FWKI 2026 10 --report
Colab/Jupyter:  main(["FWKI", "2026", "10", "--report"])     or    res = analyse("FWKI", 2026, 10)

Requires: requests, pandas, matplotlib; python-docx for the Word report.
"""
# %% Imports
import calendar
import re
import sys
import time
from pathlib import Path

import pandas as pd
import requests

OGIMET_URL = "https://www.ogimet.com/cgi-bin/getmetar"
MPS_TO_KT = 1.94384
LEVEL_SCORE = {"significant": 1, "extreme": 2, "very_extreme": 3}
LEVEL_COLOR = {"significant": "orange", "extreme": "red", "very_extreme": "darkred"}

# %% Criteria - EDIT THRESHOLDS HERE (re-run this cell and the ones below after changing)
# ---- A) LOCAL CONDITIONS ----------------------------------------------------------------
LOCAL = dict(
    min_event_tier="Severe",   # tiers at/above this are listed as events ("Moderate" lists everything)
    wind_vext_kt=34,           # mean wind or gust >= this          -> Very Extreme
    gust_spread_kt=10,         # gust minus mean wind >= this       -> Severe
    shift_deg=60,              # wind direction change between consecutive reports -> Severe
    shift_min_kt=10,           #   ...when mean speed before/after is at least this
    vis_moderate_m=5000,       # visibility 1000-5000 m             -> Moderate
    vis_severe_m=1000,         # visibility below 1000 m            -> Severe
    vis_vext_m=200,            # visibility 200 m or less           -> Very Extreme
    ceil_moderate_ft=1500,     # ceiling 500-1500 ft                -> Moderate
    ceil_severe_ft=500,        # ceiling below 500 ft               -> Severe
    ceil_vext_ft=100,          # ceiling 100 ft or less, or VV      -> Very Extreme
    heat_moderate_c=30,        # temperature >= this                -> Moderate
    heat_vext_c=35,            # temperature >= this                -> Very Extreme (extreme heat)
)
# ---- B) ICAO ANNEX 3 (SPECI-type criteria). Bands: (threshold, tag, level) -----------------
ANNEX = dict(
    vis_bands=[(1000, "VIS_LT_1000", "very_extreme"), (3000, "VIS_LT_3000", "extreme"),
               (5000, "VIS_LT_5000", "significant")],                       # visibility below x m
    cig_bands=[(100, "CIG_LE_100FT", "extreme"), (200, "CIG_LE_200FT", "extreme"),
               (500, "CIG_LE_500FT", "significant"), (1000, "CIG_LE_1000FT", "significant"),
               (1500, "CIG_LE_1500FT", "significant")],                     # ceiling at/below x ft
    dir_change_deg=60, dir_change_min_kt=10,   # wind direction change, mean speed before/after
    speed_change_kt=10,                        # mean speed change
    gust_increase_kt=10, gust_min_mean_kt=15,  # gust-spread increase, with mean speed at least this
)

TIERS = ["Normal", "Moderate", "Severe", "Very Extreme"]
TIER_LEVEL = {"Moderate": "significant", "Severe": "extreme", "Very Extreme": "very_extreme"}
CATS = {"WIND": "Wind and gusts", "PRECIP": "Rain / convective intensity", "VIS": "Visibility",
        "CEIL": "Cloud ceiling", "PHENOM": "Dangerous phenomena", "HEAT": "Heat"}
LOCAL_LABELS = {"significant": "Moderate", "extreme": "Severe", "very_extreme": "Very extreme"}
ANNEX_LABELS = {"significant": "Significant", "extreme": "Extreme", "very_extreme": "Very extreme"}
_tag = lambda cat, tier: f"{cat}_{tier.upper().replace(' ', '_')}"
LOCAL_TAG_LEVEL = {_tag(c, t): lv for c in CATS for t, lv in TIER_LEVEL.items()}


def local_info(L):
    return {
        "WIND_SEVERE": f"Gust exceeds the mean wind by {L['gust_spread_kt']} kt or more, or mean wind direction "
                       f"changed by {L['shift_deg']} degrees or more between consecutive reports "
                       f"(mean speed {L['shift_min_kt']} kt or more).",
        "WIND_VERY_EXTREME": f"Mean wind or gust of {L['wind_vext_kt']} kt or more (or P99KT).",
        "PRECIP_MODERATE": "Rain or showers without an intensity sign (RA/SHRA/TSRA), or a thunderstorm.",
        "PRECIP_SEVERE": "Heavy rain (+RA, +SHRA, +TSRA) or hail (GR).",
        "PRECIP_VERY_EXTREME": "Squall (SQ) or funnel cloud / tornado (FC).",
        "VIS_MODERATE": f"Visibility from {L['vis_severe_m']} to {L['vis_moderate_m']} m.",
        "VIS_SEVERE": f"Visibility below {L['vis_severe_m']} m.",
        "VIS_VERY_EXTREME": f"Visibility of {L['vis_vext_m']} m or less.",
        "CEIL_MODERATE": f"Ceiling (BKN/OVC) from {L['ceil_severe_ft']} to {L['ceil_moderate_ft']} ft.",
        "CEIL_SEVERE": f"Ceiling below {L['ceil_severe_ft']} ft.",
        "CEIL_VERY_EXTREME": f"Vertical visibility (VV) reported, or ceiling of {L['ceil_vext_ft']} ft or less.",
        "PHENOM_MODERATE": "Cumulonimbus (CB), towering cumulus (TCU), or thunderstorm in the vicinity (VCTS).",
        "PHENOM_SEVERE": "Thunderstorm at the aerodrome (TS) or hail (GR).",
        "PHENOM_VERY_EXTREME": "Funnel cloud / tornado (FC) or squall (SQ).",
        "HEAT_MODERATE": f"Temperature of {L['heat_moderate_c']} °C or more.",
        "HEAT_VERY_EXTREME": f"Temperature of {L['heat_vext_c']} °C or more (extreme heat).",
    }


def annex_tag_level(A):
    d = {t: lv for _, t, lv in A["vis_bands"] + A["cig_bands"]}
    d.update({"THUNDERSTORM": "extreme", "HEAVY_PRECIP": "extreme", "HAIL": "extreme",
              "FUNNEL_CLOUD": "extreme", "SQUALL": "extreme", "DUST_SAND_STORM": "extreme",
              "FREEZING_PRECIP": "extreme", "MOD_PRECIP": "significant", "FOG": "significant",
              "WIND_SHIFT": "significant", "WIND_SPEED_CHANGE": "significant",
              "GUST_INCREASE": "significant"})
    return d


def annex_info(A):
    d = {t: f"Prevailing visibility below {thr} m." for thr, t, _ in A["vis_bands"]}
    d.update({t: f"Cloud ceiling (BKN/OVC/VV) at or below {thr} ft." for thr, t, _ in A["cig_bands"]})
    d.update({
        "THUNDERSTORM": "Thunderstorm at the aerodrome (TS in present weather).",
        "HEAVY_PRECIP": "Heavy rain or drizzle (+RA, +TSRA, +SHRA, +DZ).",
        "MOD_PRECIP": "Moderate precipitation (RA/DZ etc. without an intensity sign).",
        "HAIL": "Hail or small hail (GR/GS).",
        "FUNNEL_CLOUD": "Funnel cloud, tornado or waterspout (FC).",
        "SQUALL": "Squall (SQ).",
        "DUST_SAND_STORM": "Duststorm or sandstorm (DS/SS).",
        "FREEZING_PRECIP": "Freezing rain or drizzle (FZRA/FZDZ).",
        "FOG": "Fog (FG, excluding shallow MIFG).",
        "WIND_SHIFT": f"Mean wind direction changed by {A['dir_change_deg']} degrees or more between consecutive "
                      f"reports (mean speed {A['dir_change_min_kt']} kt or more before and/or after).",
        "WIND_SPEED_CHANGE": f"Mean wind speed changed by {A['speed_change_kt']} kt or more between consecutive reports.",
        "GUST_INCREASE": f"Gust spread (gust minus mean) rose by {A['gust_increase_kt']} kt or more between "
                         f"consecutive reports, with mean speed of {A['gust_min_mean_kt']} kt or more.",
    })
    return d


# %% Download from Ogimet
def fetch_metars(icao, year, month, cache_dir="cache", retries=3):
    cache = Path(cache_dir)
    cache.mkdir(exist_ok=True)
    path = cache / f"{icao}_{year}{month:02d}.txt"
    now = pd.Timestamp.now("UTC")
    complete = (year, month) < (now.year, now.month)       # only finished months are cached
    if complete and path.exists() and path.stat().st_size > 0:
        return path.read_text()
    last_day = calendar.monthrange(year, month)[1]
    params = {"icao": icao, "begin": f"{year}{month:02d}010000",
              "end": f"{year}{month:02d}{last_day}2359"}
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(OGIMET_URL, params=params, timeout=90,
                             headers={"User-Agent": "Mozilla/5.0"})
            if r.ok and r.text.strip():
                if complete:
                    path.write_text(r.text)
                return r.text
            print(f"[warn] attempt {attempt}: HTTP {r.status_code} / empty reply")
        except requests.RequestException as e:
            print(f"[warn] attempt {attempt}: {e}")
        time.sleep(5 * attempt)
    sys.exit("Could not download data from Ogimet. Try again in a few minutes.")


# %% Parse METARs
LINE_RE = re.compile(r"^(\d{12})\s+(METAR|SPECI)\s+(.*?)\s*=?\s*$")
CSV_RE = re.compile(r"^\w{4},(\d{4}),(\d{2}),(\d{2}),(\d{2}),(\d{2}),(METAR|SPECI)\s+(.*?)\s*=?\s*$")
WIND_RE = re.compile(r"^(\d{3}|VRB)(\d{2,3})(?:G(\d{2,3}))?(KT|MPS)$")
TEMP_RE = re.compile(r"^(M?\d{2})/(M?\d{2})?$")
QNH_RE = re.compile(r"^Q(\d{4})$")
ALT_RE = re.compile(r"^A(\d{4})$")
CLD_RE = re.compile(r"^(FEW|SCT|BKN|OVC|VV)(\d{3})(CB|TCU)?$")
VIS_RE = re.compile(r"^\d{4}$")
WX_RE = re.compile(r"^(\+|-|VC)?(MI|BC|PR|DR|BL|SH|TS|FZ)*"
                   r"(DZ|RA|SN|SG|PL|GR|GS|UP|BR|FG|FU|VA|DU|SA|HZ|PO|SQ|FC|SS|DS)*$")
STOP = {"RMK", "TEMPO", "BECMG", "NOSIG"}
NUM_COLS = ["wind_dir", "wind_kt", "gust_kt", "vis_m", "temp_c", "dew_c", "qnh_hpa", "ceiling_ft"]


def _temp(s):
    return int(s.replace("M", "-")) if s else None


def parse_report(stamp, rtype, body):
    rec = dict(time=pd.to_datetime(stamp, format="%Y%m%d%H%M"), rtype=rtype,
               wind_dir=None, wind_kt=None, gust_kt=None, vis_m=None, temp_c=None, dew_c=None,
               qnh_hpa=None, ceiling_ft=None, vv=False, cb=False, tcu=False, wind_p99=False,
               wx="", raw=body)
    wx = []
    for t in body.split():
        if t in STOP:
            break
        if t == "CAVOK":
            rec["vis_m"] = 9999
        elif t in ("P99KT", "ABV99KT") or re.match(r"^(\d{3}|VRB)P99KT$", t):
            rec["wind_p99"] = True
        elif (m := WIND_RE.match(t)):
            k = MPS_TO_KT if m.group(4) == "MPS" else 1
            rec["wind_dir"] = None if m.group(1) == "VRB" else int(m.group(1))
            rec["wind_kt"] = round(int(m.group(2)) * k)
            rec["gust_kt"] = round(int(m.group(3)) * k) if m.group(3) else None
        elif (m := TEMP_RE.match(t)):
            rec["temp_c"], rec["dew_c"] = _temp(m.group(1)), _temp(m.group(2))
        elif (m := QNH_RE.match(t)):
            rec["qnh_hpa"] = int(m.group(1))
        elif (m := ALT_RE.match(t)):
            rec["qnh_hpa"] = round(int(m.group(1)) / 100 * 33.8639, 1)
        elif (m := CLD_RE.match(t)):
            rec["cb"] |= m.group(3) == "CB"
            rec["tcu"] |= m.group(3) == "TCU"
            if m.group(1) == "VV":
                rec["vv"] = True
            if m.group(1) in ("BKN", "OVC", "VV"):
                h = int(m.group(2)) * 100
                if rec["ceiling_ft"] is None or h < rec["ceiling_ft"]:
                    rec["ceiling_ft"] = h
        elif VIS_RE.match(t) and rec["vis_m"] is None:
            rec["vis_m"] = int(t)
        elif t and WX_RE.match(t):
            wx.append(t)
    rec["wx"] = " ".join(wx)
    return rec


def parse_metars(text):
    """Accepts Ogimet CSV lines ('FWKI,2026,10,01,07,00,METAR FWKI ...=') or plain lines.
    NIL reports are skipped."""
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if (m := CSV_RE.match(line)):
            stamp, rtype, body = "".join(m.group(i) for i in range(1, 6)), m.group(6), m.group(7)
        elif (m := LINE_RE.match(line)):
            stamp, rtype, body = m.group(1), m.group(2), m.group(3)
        else:
            continue
        if body.split()[-1] == "NIL":
            continue
        rows.append(parse_report(stamp, rtype, body))
    if not rows:
        sys.exit("No METAR lines could be parsed - delete the file in cache/ and try again.")
    df = pd.DataFrame(rows).drop_duplicates(subset=["time", "raw"]).sort_values("time")
    for c in NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["wind_peak_kt"] = df[["wind_kt", "gust_kt"]].max(axis=1)
    return df.reset_index(drop=True)


# %% Shared detection helpers
def change_tags(df, deg, deg_min_kt, spd_kt, gust_kt, gust_min_kt, max_gap_h=3):
    """Changes between consecutive reports: direction shift, speed change, gust-spread increase."""
    p = df.shift(1)
    ok = (df["time"] - p["time"]) <= pd.Timedelta(hours=max_gap_h)
    ddir = ((df["wind_dir"] - p["wind_dir"] + 180) % 360 - 180).abs()
    vmax = pd.concat([df["wind_kt"], p["wind_kt"]], axis=1).max(axis=1)
    shift = ok & (ddir >= deg) & (vmax >= deg_min_kt)
    spd = ok & ((df["wind_kt"] - p["wind_kt"]).abs() >= spd_kt)
    g_now = (df["gust_kt"] - df["wind_kt"]).fillna(0)
    g_prev = (p["gust_kt"] - p["wind_kt"]).fillna(0)
    gust = ok & ((g_now - g_prev) >= gust_kt) & (vmax >= gust_min_kt)
    return shift, spd, gust


def peak_spec(etype):
    if etype in ("WIND_SHIFT", "WIND_SPEED_CHANGE", "GUST_INCREASE"):
        return None
    for prefix, spec in (("WIND_", ("wind_peak_kt", "max")), ("VIS_", ("vis_m", "min")),
                         ("CEIL_", ("ceiling_ft", "min")), ("CIG_", ("ceiling_ft", "min")),
                         ("HEAT_", ("temp_c", "max"))):
        if etype.startswith(prefix):
            return spec
    return None


def explode_events(df, tags, tag_level):
    d = df.copy()
    d["events"] = tags
    ev = d.explode("events").dropna(subset=["events"]).rename(columns={"events": "event"})
    ev["level"] = ev["event"].map(tag_level)
    return ev.reset_index(drop=True)


def build_episodes(ev, tag_level, gap_hours):
    rows = []
    for etype, g in ev.groupby("event"):
        g = g.sort_values("time")
        g = g.assign(ep=(g["time"].diff() > pd.Timedelta(hours=gap_hours)).cumsum())
        for _, e in g.groupby("ep"):
            row = dict(event=etype, level=tag_level[etype], start=e["time"].min(),
                       end=e["time"].max(), reports=len(e))
            spec = peak_spec(etype)
            if spec:
                row["peak"], row["peak_var"] = getattr(e[spec[0]], spec[1])(), spec[0]
            rows.append(row)
    ep = pd.DataFrame(rows, columns=["event", "level", "start", "end", "reports", "peak", "peak_var"])
    return ep.sort_values("start").reset_index(drop=True) if len(ep) else ep


# %% A) Local conditions - tiered classification
def local_row_tiers(r, shift, L):
    wx = r["wx"].split()
    mean = r["wind_kt"] if pd.notna(r["wind_kt"]) else 0
    gust = r["gust_kt"] if pd.notna(r["gust_kt"]) else 0
    ts = any("TS" in w and not w.startswith("VC") for w in wx)
    vcts = any(w.startswith("VCTS") for w in wx)
    hail = any("GR" in w for w in wx)
    sq = any(w.lstrip("+-").startswith("SQ") for w in wx)
    fc = any("FC" in w for w in wx)
    heavy = any(w.startswith("+") and "RA" in w for w in wx)
    rain = any("RA" in w and not w.startswith(("-", "VC")) for w in wx)

    wind = "Normal"
    if (gust > 0 and gust - mean >= L["gust_spread_kt"]) or shift:
        wind = "Severe"
    if mean >= L["wind_vext_kt"] or gust >= L["wind_vext_kt"] or r["wind_p99"]:
        wind = "Very Extreme"

    precip = "Normal"
    if rain or ts:
        precip = "Moderate"
    if heavy or hail:
        precip = "Severe"
    if sq or fc:
        precip = "Very Extreme"

    vis, v = "Normal", r["vis_m"]
    if pd.notna(v):
        if v <= L["vis_vext_m"]:
            vis = "Very Extreme"
        elif v < L["vis_severe_m"]:
            vis = "Severe"
        elif v <= L["vis_moderate_m"]:
            vis = "Moderate"

    ceil, c = "Normal", r["ceiling_ft"]
    if r["vv"] or (pd.notna(c) and c <= L["ceil_vext_ft"]):
        ceil = "Very Extreme"
    elif pd.notna(c) and c < L["ceil_severe_ft"]:
        ceil = "Severe"
    elif pd.notna(c) and c <= L["ceil_moderate_ft"]:
        ceil = "Moderate"

    phen = "Normal"
    if r["cb"] or r["tcu"] or vcts:
        phen = "Moderate"
    if ts or hail:
        phen = "Severe"
    if fc or sq:
        phen = "Very Extreme"

    heat, t = "Normal", r["temp_c"]
    if pd.notna(t):
        if t >= L["heat_vext_c"]:
            heat = "Very Extreme"
        elif t >= L["heat_moderate_c"]:
            heat = "Moderate"
    return dict(WIND=wind, PRECIP=precip, VIS=vis, CEIL=ceil, PHENOM=phen, HEAT=heat)


def detect_local(df, gap_hours=3, L=None):
    L = L or LOCAL
    df = df.copy()
    shift, _, _ = change_tags(df, L["shift_deg"], L["shift_min_kt"], 1e9, 1e9, 1e9)
    tiers = [local_row_tiers(r, bool(shift.iloc[i]), L) for i, r in enumerate(df.to_dict("records"))]
    for cat in CATS:
        df[f"{cat.lower()}_tier"] = [t[cat] for t in tiers]
    df["overall_tier"] = [max((t[c] for c in CATS), key=TIERS.index) for t in tiers]
    min_i = TIERS.index(L["min_event_tier"])
    tags = [[_tag(c, t[c]) for c in CATS if TIERS.index(t[c]) >= max(min_i, 1)] for t in tiers]
    ev = explode_events(df, tags, LOCAL_TAG_LEVEL)
    sec = dict(key="local", title="Local conditions (tiered criteria)", ev=ev,
               ep=build_episodes(ev, LOCAL_TAG_LEVEL, gap_hours), tag_level=LOCAL_TAG_LEVEL,
               info=local_info(L), labels=LOCAL_LABELS,
               description=("Each report is classified Normal / Moderate / Severe / Very Extreme for wind, rain, "
                            "visibility, cloud ceiling, dangerous phenomena and heat. Reports at "
                            f"'{L['min_event_tier']}' or above are listed as events; all tiers are counted "
                            "in the tier table."))
    return df, sec


# %% B) ICAO Annex 3 criteria
def band(value, bands, inclusive):
    for thr, tag, _ in bands:
        if (value <= thr) if inclusive else (value < thr):
            return tag
    return None


def tag_report_annex(r, A):
    tags, wx = [], r["wx"].split()
    if pd.notna(r["vis_m"]) and (t := band(r["vis_m"], A["vis_bands"], False)):
        tags.append(t)
    if pd.notna(r["ceiling_ft"]) and (t := band(r["ceiling_ft"], A["cig_bands"], True)):
        tags.append(t)
    if any("TS" in w and not w.startswith("VC") for w in wx):
        tags.append("THUNDERSTORM")
    if any(w.startswith("+") and re.search("RA|DZ", w) for w in wx):
        tags.append("HEAVY_PRECIP")
    elif any(re.search("RA|DZ|SN|SG|PL|UP", w) and not w.startswith(("-", "+", "VC")) for w in wx):
        tags.append("MOD_PRECIP")
    if any(re.search("GR|GS", w) for w in wx):
        tags.append("HAIL")
    if any("FC" in w for w in wx):
        tags.append("FUNNEL_CLOUD")
    if any(w.lstrip("+-").startswith("SQ") for w in wx):
        tags.append("SQUALL")
    if any(re.search("DS|SS", w) for w in wx):
        tags.append("DUST_SAND_STORM")
    if any("FZ" in w and re.search("RA|DZ", w) for w in wx):
        tags.append("FREEZING_PRECIP")
    if any("FG" in w and not w.startswith("MI") for w in wx):
        tags.append("FOG")
    return tags


def detect_annex(df, gap_hours=3, A=None):
    A = A or ANNEX
    tags = [tag_report_annex(r, A) for r in df.to_dict("records")]
    shift, spd, gust = change_tags(df, A["dir_change_deg"], A["dir_change_min_kt"], A["speed_change_kt"],
                                   A["gust_increase_kt"], A["gust_min_mean_kt"])
    for i in range(len(df)):
        if shift.iloc[i]:
            tags[i].append("WIND_SHIFT")
        if spd.iloc[i]:
            tags[i].append("WIND_SPEED_CHANGE")
        if gust.iloc[i]:
            tags[i].append("GUST_INCREASE")
    tl = annex_tag_level(A)
    ev = explode_events(df, tags, tl)
    return dict(key="annex3", title="ICAO Annex 3 criteria", ev=ev,
                ep=build_episodes(ev, tl, gap_hours), tag_level=tl, info=annex_info(A),
                labels=ANNEX_LABELS,
                description=("Visibility and ceiling bands, onset of significant weather phenomena and "
                             "wind-change criteria of the kind used to trigger SPECI reports under ICAO Annex 3. "
                             "The visibility bands have been customised (see criteria table)."))


# %% Run the analysis and print a summary
def analyse(icao, year, month, gap_hours=3, local=None, annex=None):
    """local / annex: optional dicts overriding entries of LOCAL / ANNEX for this run only."""
    icao = icao.upper()
    text = fetch_metars(icao, year, month)
    try:
        df = parse_metars(text)
    except SystemExit:
        (Path("cache") / f"{icao}_{year}{month:02d}.txt").unlink(missing_ok=True)   # drop bad download
        raise
    df, loc_sec = detect_local(df, gap_hours, {**LOCAL, **(local or {})})
    annex_sec = detect_annex(df, gap_hours, {**ANNEX, **(annex or {})})
    return dict(icao=icao, year=year, month=month, df=df, sections=[loc_sec, annex_sec])


def summarise(res):
    df, icao, year, month = res["df"], res["icao"], res["year"], res["month"]
    days = calendar.monthrange(year, month)[1]
    print(f"\n=== {icao}  {year}-{month:02d} ===")
    print(f"Reports: {len(df)} (≈{len(df) / (days * 24) * 100:.0f}% of hourly expected), "
          f"of which SPECI: {(df['rtype'] == 'SPECI').sum()}")
    for label, col, fn, unit in (("Max temp", "temp_c", "idxmax", "°C"), ("Min temp", "temp_c", "idxmin", "°C"),
                                 ("Peak wind/gust", "wind_peak_kt", "idxmax", "kt"),
                                 ("Lowest visibility", "vis_m", "idxmin", "m"),
                                 ("Lowest ceiling", "ceiling_ft", "idxmin", "ft")):
        if df[col].notna().any():
            i = getattr(df[col], fn)()
            print(f"{label}: {df.at[i, col]:.0f} {unit} at {df.at[i, 'time']}")
    print("\nLOCAL tiers (number of reports):")
    print(pd.DataFrame({CATS[c]: df[f"{c.lower()}_tier"].value_counts() for c in CATS})
          .reindex(TIERS).fillna(0).astype(int).T.to_string())
    for sec in res["sections"]:
        print(f"\n--- {sec['title']}: episodes ---")
        ep = sec["ep"]
        if ep.empty:
            print("  none detected")
        else:
            print(ep.assign(level=ep["level"].map(sec["labels"])).groupby(["level", "event"])
                  .size().rename("episodes").to_string())


# %% Graphs (values with shaded events, event timeline, day x hour heatmap)
def panel_of(n):
    if n.startswith("HEAT_"):
        return 0
    if n.startswith(("VIS_", "FOG")):
        return 2
    if n.startswith(("CEIL_", "CIG_")):
        return 3
    return 1


def _finish(fig, out):
    import matplotlib.pyplot as plt
    fig.tight_layout()
    if out:
        fig.savefig(out, dpi=150)
        plt.close(fig)
        print(f"Plot saved: {out}")
    else:
        plt.show()


def make_plot(res, sec, out=None):
    import matplotlib.pyplot as plt
    df, ep = res["df"], sec["ep"]
    fig, ax = plt.subplots(4, 1, figsize=(11, 10), sharex=True)
    ax[0].plot(df["time"], df["temp_c"], lw=1, color="tab:red"); ax[0].set_ylabel("Temp (°C)")
    ax[1].plot(df["time"], df["wind_peak_kt"], lw=1, color="tab:blue"); ax[1].set_ylabel("Wind/gust (kt)")
    ax[2].plot(df["time"], df["vis_m"], lw=1, color="tab:gray"); ax[2].set_ylabel("Visibility (m)")
    ax[3].plot(df["time"], df["ceiling_ft"], ".", ms=3, color="tab:green"); ax[3].set_ylabel("Ceiling (ft)")
    for _, e in ep.iterrows():
        ax[panel_of(e["event"])].axvspan(e["start"], e["end"] + pd.Timedelta(hours=1), alpha=0.35,
                                         color=LEVEL_COLOR[e["level"]])
    lab = sec["labels"]
    fig.suptitle(f"{res['icao']} {res['year']}-{res['month']:02d} - {sec['title']}\n"
                 f"(dark red = {lab['very_extreme']}, red = {lab['extreme']}, orange = {lab['significant']})")
    _finish(fig, out)


def plot_timeline(res, sec, out=None):
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt
    ev, year, month = sec["ev"], res["year"], res["month"]
    if ev.empty:
        print("No events to plot.")
        return
    order = (ev.assign(s=ev["level"].map(LEVEL_SCORE)).groupby("event")["s"].max()
             .sort_values().index.tolist())
    ypos = {e: i for i, e in enumerate(order)}
    fig, ax = plt.subplots(figsize=(12, 0.55 * len(order) + 2.2))
    for level in ("significant", "extreme", "very_extreme"):
        g = ev[ev["level"] == level]
        if not g.empty:
            ax.scatter(g["time"], g["event"].map(ypos), marker="|", s=450, linewidths=2.2,
                       color=LEVEL_COLOR[level], label=sec["labels"][level])
    start = pd.Timestamp(year, month, 1)
    ax.set_xlim(start, start + pd.offsets.MonthEnd(0) + pd.Timedelta(days=1))
    ax.set_yticks(range(len(order))); ax.set_yticklabels(order); ax.set_ylim(-0.7, len(order) - 0.3)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d"))
    ax.tick_params(axis="x", labelsize=7)
    ax.grid(axis="x", alpha=0.3)
    ax.set_xlabel(f"Day of month, {year}-{month:02d} (UTC)")
    ax.set_title(f"{res['icao']} - {sec['title']}: when each event occurred")
    ax.legend(loc="upper right", ncol=3, fontsize=8, framealpha=0.9)
    _finish(fig, out)


def plot_heatmap(res, sec, out=None):
    """Day x hour grid. White = no report, grey = report but nothing flagged, colours = worst event."""
    import numpy as np
    import matplotlib.pyplot as plt
    from matplotlib.colors import ListedColormap
    from matplotlib.patches import Patch
    df, ev, year, month = res["df"], sec["ev"], res["year"], res["month"]
    days = calendar.monthrange(year, month)[1]
    grid = np.zeros((24, days))
    for t in df["time"]:
        grid[t.hour, t.day - 1] = max(grid[t.hour, t.day - 1], 1)
    for _, r in ev.iterrows():
        t = r["time"]
        grid[t.hour, t.day - 1] = max(grid[t.hour, t.day - 1], 1 + LEVEL_SCORE[r["level"]])
    colors = ["white", "#dcdcdc", "orange", "red", "darkred"]
    fig, ax = plt.subplots(figsize=(12, 6.5))
    ax.imshow(grid, aspect="auto", origin="lower", cmap=ListedColormap(colors), vmin=0, vmax=4,
              interpolation="nearest")
    ax.set_xticks(range(days)); ax.set_xticklabels(range(1, days + 1), fontsize=7)
    ax.set_yticks(range(0, 24, 2)); ax.set_yticklabels([f"{h:02d}" for h in range(0, 24, 2)], fontsize=8)
    ax.set_xticks(np.arange(-0.5, days, 1), minor=True); ax.set_yticks(np.arange(-0.5, 24, 1), minor=True)
    ax.grid(which="minor", color="white", lw=0.8); ax.tick_params(which="minor", length=0)
    ax.set_xlabel(f"Day of month, {year}-{month:02d}"); ax.set_ylabel("Hour (UTC)")
    ax.set_title(f"{res['icao']} - {sec['title']}: worst event by day and hour")
    names = ["no report / NIL", "report, nothing flagged", sec["labels"]["significant"],
             sec["labels"]["extreme"], sec["labels"]["very_extreme"]]
    ax.legend(handles=[Patch(facecolor=c, edgecolor="grey", label=n) for c, n in zip(colors, names)],
              loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=5, fontsize=8, frameon=False)
    _finish(fig, out)


# %% Word report
UNITS = {"vis_m": "m", "ceiling_ft": "ft", "wind_peak_kt": "kt", "temp_c": "°C"}
LEVEL_FILL = {"significant": "FFD966", "extreme": "FF8F8F", "very_extreme": "C00000"}
LOGO_FILE = "dccms_logo.png"


def find_logo():
    """Return the path of dccms_logo.png (next to this script or in the current folder), else None."""
    dirs = [Path.cwd()]
    try:
        dirs.insert(0, Path(__file__).resolve().parent)
    except NameError:                       # pasted into Colab/Jupyter: no __file__
        pass
    for d in dirs:
        if (d / LOGO_FILE).exists():
            return d / LOGO_FILE
    return None


def write_report(res, outdir="output"):
    """Detailed Word report with one part for Local conditions and one for ICAO Annex 3."""
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    icao, year, month, df = res["icao"], res["year"], res["month"], res["df"]
    outdir = Path(outdir); outdir.mkdir(exist_ok=True)
    stem = f"{icao}_{year}{month:02d}"
    mname = f"{calendar.month_name[month]} {year}"
    fmt = lambda t: t.strftime("%d %b %H:%M") + "Z"

    doc = Document()
    sec0 = doc.sections[0]
    sec0.page_width, sec0.page_height = Cm(21), Cm(29.7)
    sec0.left_margin = sec0.right_margin = sec0.top_margin = sec0.bottom_margin = Cm(2)
    logo = find_logo()
    if logo:
        sec0.top_margin, sec0.header_distance = Cm(2.8), Cm(0.8)
        sec0.header.paragraphs[0].add_run().add_picture(str(logo), width=Cm(6.0))
    else:
        print(f"[note] {LOGO_FILE} not found - report written without the DCCMS logo.")
    doc.styles["Normal"].font.name, doc.styles["Normal"].font.size = "Calibri", Pt(10.5)
    for name, size in (("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11)):
        doc.styles[name].font.name, doc.styles[name].font.size = "Calibri", Pt(size)
        doc.styles[name].font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

    fp = sec0.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.add_run(f"{icao} extreme weather report - {mname}   |   Page ").font.size = Pt(8)
    r2 = fp.add_run(); r2.font.size = Pt(8)
    for kind, text in (("begin", None), (None, " PAGE "), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = text
        r2._r.append(el)

    def shade(cell, fill):
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
        cell._tc.get_or_add_tcPr().append(shd)

    def put(cell, text, bold=False, size=9, color=None):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(text)); r.bold, r.font.size = bold, Pt(size)
        if color:
            r.font.color.rgb = RGBColor.from_string(color)

    def table(header, rows, widths, level_col=None, labels=None, size=9):
        t = doc.add_table(rows=1, cols=len(header))
        t.style, t.alignment, t.autofit = "Table Grid", WD_TABLE_ALIGNMENT.CENTER, False
        for i, h in enumerate(header):
            put(t.rows[0].cells[i], h, True, size, "FFFFFF"); shade(t.rows[0].cells[i], "1F3A5F")
        hdr = OxmlElement("w:tblHeader"); hdr.set(qn("w:val"), "true")
        t.rows[0]._tr.get_or_add_trPr().append(hdr)
        for row in rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                put(cells[i], v, size=size)
                if level_col is not None and i == level_col and labels:
                    lv = next((k for k, lab in labels.items() if lab == v), None)
                    if lv:
                        shade(cells[i], LEVEL_FILL[lv])
                        put(cells[i], v, True, size, "FFFFFF" if lv == "very_extreme" else None)
        for i, w in enumerate(widths):
            t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ---- title and overview ----
    p = doc.add_paragraph()
    r = p.add_run("Extreme Weather Events Report"); r.bold, r.font.size = True, Pt(24)
    r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    r = doc.add_paragraph().add_run(f"Aerodrome {icao}  |  {mname}")
    r.font.size, r.font.color.rgb = Pt(14), RGBColor(0x59, 0x59, 0x59)
    r = doc.add_paragraph().add_run(
        "Source: hourly METAR/SPECI reports from Ogimet. All times UTC (Malawi local time = UTC+2). "
        f"Generated {pd.Timestamp.now().strftime('%d %b %Y %H:%M')}.")
    r.italic, r.font.size = True, Pt(9)

    doc.add_heading("1. Overview", level=1)
    exp = calendar.monthrange(year, month)[1] * 24
    doc.add_paragraph(
        f"This report analyses {len(df)} METAR/SPECI reports for {icao} between {fmt(df['time'].min())} and "
        f"{fmt(df['time'].max())} ({(df['rtype'] == 'SPECI').sum()} SPECI). Coverage is about "
        f"{len(df) / exp * 100:.0f}% of the hourly reports expected for the month; NIL reports are excluded. "
        "The same reports are assessed twice: Part A uses local tiered criteria, Part B uses ICAO Annex 3 criteria.")
    rows = []
    for label, col, fn, unit in (("Highest temperature", "temp_c", "idxmax", "°C"),
                                 ("Lowest temperature", "temp_c", "idxmin", "°C"),
                                 ("Strongest wind / gust", "wind_peak_kt", "idxmax", "kt"),
                                 ("Lowest visibility", "vis_m", "idxmin", "m"),
                                 ("Lowest cloud ceiling", "ceiling_ft", "idxmin", "ft"),
                                 ("Lowest QNH", "qnh_hpa", "idxmin", "hPa"),
                                 ("Highest QNH", "qnh_hpa", "idxmax", "hPa")):
        if df[col].notna().any():
            i = getattr(df[col], fn)()
            rows.append((label, f"{df.at[i, col]:.0f} {unit}", fmt(df.at[i, "time"])))
    table(["Parameter", "Value", "Time (UTC)"], rows, [6, 5, 6])
    comp = [(s["title"], len(s["ep"]), int((s["ep"]["level"] == "very_extreme").sum()),
             int((s["ep"]["level"] == "extreme").sum()), int((s["ep"]["level"] == "significant").sum()))
            for s in res["sections"]]
    table(["Criteria set", "Episodes", "Top level", "Middle level", "Lowest level"], comp, [7, 2.2, 2.6, 2.6, 2.6])

    # ---- one part per criteria set ----
    for letter, sec in zip("AB", res["sections"]):
        ev, ep, labels, tl = sec["ev"].copy(), sec["ep"], sec["labels"], sec["tag_level"]
        ev["score"] = ev["level"].map(LEVEL_SCORE)
        doc.add_heading(f"Part {letter} - {sec['title']}", level=1)
        doc.add_paragraph(sec["description"])

        stats = []
        for etype, g in ev.groupby("event"):
            stats.append(dict(event=etype, level=tl[etype], score=LEVEL_SCORE[tl[etype]],
                              episodes=int((ep["event"] == etype).sum()), reports=len(g),
                              first=g["time"].min(), last=g["time"].max()))
        stats.sort(key=lambda d: (-d["score"], -d["episodes"], d["event"]))

        # narrative
        if stats:
            by = ep.groupby("level").size().to_dict()
            top = max(stats, key=lambda d: (d["episodes"], d["score"]))
            dd = ev.assign(day=ev["time"].dt.date).groupby("day").agg(s=("score", "max"), n=("event", "size"))
            wd = dd.sort_values(["s", "n"], ascending=False).index[0]
            lg = ep.assign(dur=ep["end"] - ep["start"]).sort_values("dur", ascending=False).iloc[0]
            doc.add_paragraph(
                f"{len(ep)} episodes were identified: {by.get('very_extreme', 0)} {labels['very_extreme'].lower()}, "
                f"{by.get('extreme', 0)} {labels['extreme'].lower()} and {by.get('significant', 0)} "
                f"{labels['significant'].lower()}. The most frequent event was {top['event']} "
                f"({top['episodes']} episodes). The most affected day was {wd.strftime('%d %B')} with "
                f"{int(dd.loc[wd, 'n'])} flagged reports. The longest episode was {lg['event']}, "
                f"{fmt(lg['start'])} to {fmt(lg['end'])}.")
        else:
            doc.add_paragraph("No events were detected with these criteria.")

        doc.add_heading(f"{letter}.1 Criteria used", level=2)
        if sec["key"] == "local":
            trows = []
            for c, name in CATS.items():
                counts = df[f"{c.lower()}_tier"].value_counts()
                trows.append((name, *[int(counts.get(t, 0)) for t in TIERS]))
            doc.add_paragraph("Number of reports in each tier:")
            table(["Category", *TIERS], trows, [5.5, 2.5, 2.5, 2.5, 3.0])
        table(["Event", "Level", "Definition"],
              [(e, labels[tl[e]], sec["info"][e]) for e in sorted(sec["info"], key=lambda k: (-LEVEL_SCORE[tl[k]], k))],
              [3.8, 2.4, 10.8], level_col=1, labels=labels, size=8.5)

        doc.add_heading(f"{letter}.2 Summary of events", level=2)
        if stats:
            table(["Event", "Level", "Episodes", "Reports", "First", "Last"],
                  [(d["event"], labels[d["level"]], d["episodes"], d["reports"], fmt(d["first"]), fmt(d["last"]))
                   for d in stats], [4.0, 2.4, 1.8, 1.8, 3.5, 3.5], level_col=1, labels=labels)
        else:
            doc.add_paragraph("No events detected.")

        figs = [(outdir / f"{stem}_{sec['key']}.png", "Values with event periods shaded."),
                (outdir / f"{stem}_{sec['key']}_timeline.png", "Timeline of every flagged report by event type."),
                (outdir / f"{stem}_{sec['key']}_heatmap.png", "Day-by-hour map of the worst event in each hour.")]
        figs = [f for f in figs if f[0].exists()]
        if figs:
            doc.add_heading(f"{letter}.3 Figures", level=2)
            for k, (path, cap) in enumerate(figs, 1):
                doc.add_picture(str(path), width=Cm(16.0))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                c = doc.add_paragraph(f"Figure {letter}{k}. {cap}")
                c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                c.runs[0].italic, c.runs[0].font.size = True, Pt(9)

        doc.add_heading(f"{letter}.4 Daily summary", level=2)
        if ev.empty:
            doc.add_paragraph("No flagged days.")
        else:
            drows = []
            for day, g in ev.assign(day=ev["time"].dt.date).groupby("day"):
                worst = g.sort_values("score", ascending=False).iloc[0]["level"]
                names = g.drop_duplicates("event").sort_values("score", ascending=False)["event"].tolist()
                drows.append((day.strftime("%a %d %b"), labels[worst], len(g), ", ".join(names)))
            table(["Date", "Worst level", "Flags", "Events"], drows, [2.8, 2.6, 1.4, 10.2],
                  level_col=1, labels=labels, size=8.5)

        doc.add_heading(f"{letter}.5 Detailed event log", level=2)
        doc.add_paragraph("Each episode groups consecutive flagged reports (gap of up to 3 hours). "
                          "'Report at peak' is the raw METAR/SPECI that gave the peak value, or the first "
                          "report of the episode when the event has no measured peak.")
        for d in stats:
            etype = d["event"]
            doc.add_heading(f"{etype.replace('_', ' ')}  -  {labels[d['level']]}", level=3)
            doc.add_paragraph(sec["info"][etype]).runs[0].italic = True
            doc.add_paragraph(f"{d['episodes']} episode(s), {d['reports']} flagged report(s). "
                              f"First: {fmt(d['first'])}. Last: {fmt(d['last'])}.")
            rows = []
            for _, e in ep[ep["event"] == etype].sort_values("start").iterrows():
                sub = ev[(ev["event"] == etype) & (ev["time"] >= e["start"]) & (ev["time"] <= e["end"])]
                sub = sub.reset_index(drop=True)
                spec, pk, k = peak_spec(etype), "-", 0
                if spec and sub[spec[0]].notna().any():
                    k = sub[spec[0]].idxmax() if spec[1] == "max" else sub[spec[0]].idxmin()
                    pk = f"{sub.at[k, spec[0]]:.0f} {UNITS.get(spec[0], '')}"
                hrs = (e["end"] - e["start"]).total_seconds() / 3600
                rows.append((fmt(e["start"]), fmt(e["end"]),
                             "single report" if e["reports"] == 1 else f"{hrs:.0f} h",
                             e["reports"], pk, sub.at[k, "raw"]))
            table(["Start", "End", "Duration", "Reports", "Peak", "Report at peak / first report"],
                  rows, [2.5, 2.5, 1.9, 1.3, 1.5, 7.3], size=8)

    doc.add_heading("Notes and limitations", level=1)
    for txt in ("Results depend on report availability; missing or NIL reports can hide events.",
                "Only prevailing visibility and the lowest BKN/OVC/VV layer are used; RVR and runway-specific "
                "data are not analysed.",
                "Present-weather flags use the main body of each report; remarks (RMK) and forecast trend "
                "groups (TEMPO/BECMG) are ignored.",
                "Part A thresholds are local settings and should be agreed with the national meteorological "
                "service. Part B wind-change values use the common ICAO figures; some states apply national values."):
        doc.add_paragraph(txt, style="List Bullet")

    path = outdir / f"{stem}_report.docx"
    doc.save(path)
    return path


# %% Save CSVs, graphs and report
def save_outputs(res, outdir="output", plots=True, report=True):
    out = Path(outdir); out.mkdir(exist_ok=True)
    stem = f"{res['icao']}_{res['year']}{res['month']:02d}"
    res["df"].to_csv(out / f"{stem}_parsed.csv", index=False)
    for sec in res["sections"]:
        sec["ev"].to_csv(out / f"{stem}_{sec['key']}_flagged_reports.csv", index=False)
        sec["ep"].to_csv(out / f"{stem}_{sec['key']}_episodes.csv", index=False)
        if plots or report:
            make_plot(res, sec, out / f"{stem}_{sec['key']}.png")
            plot_timeline(res, sec, out / f"{stem}_{sec['key']}_timeline.png")
            plot_heatmap(res, sec, out / f"{stem}_{sec['key']}_heatmap.png")
    if report:
        path = write_report(res, out)
        print("Report saved:", path)
        return path
    print("Saved to", out.resolve())


# %% CLI
def main(argv=None):
    import argparse
    p = argparse.ArgumentParser(description="Ogimet METAR extreme-event analysis (Local + ICAO Annex 3)")
    p.add_argument("icao"); p.add_argument("year", type=int); p.add_argument("month", type=int)
    p.add_argument("--gap", type=float, default=3, help="hours gap that splits episodes")
    p.add_argument("--outdir", default="output")
    p.add_argument("--plot", action="store_true", help="save graphs")
    p.add_argument("--report", action="store_true", help="write the Word report (needs python-docx)")
    a = p.parse_args(argv)
    if "ipykernel" not in sys.modules:
        import matplotlib
        matplotlib.use("Agg")
    res = analyse(a.icao, a.year, a.month, a.gap)
    summarise(res)
    save_outputs(res, a.outdir, plots=a.plot or a.report, report=a.report)
    return res


if __name__ == "__main__" and "ipykernel" not in sys.modules:
    main()  # command line only; in Colab/Jupyter call main([...]) or analyse(...) yourself
