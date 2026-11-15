#!/usr/bin/env python3
"""
metar_extremes.py - monthly METAR extreme-event analysis from Ogimet, with a 2-page bulletin.

One set of criteria, based on ICAO Annex 3 (SPECI-type criteria): visibility and ceiling bands,
onset of significant weather, and wind-change criteria. Every threshold can be adjusted for
local conditions in the CRITERIA dict below (or in the web app sidebar).

Command line:   python metar_extremes.py FWKI 2026 10 --report
Colab/Jupyter:  main(["FWKI", "2026", "10", "--report"])     or    res = analyse("FWKI", 2026, 10)

Requires: requests, pandas, matplotlib; python-docx for the Word bulletin.
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
LEVEL_LABELS = {"significant": "Significant", "extreme": "Extreme", "very_extreme": "Very extreme"}
STATION_NAMES = {"FWKI": "Lilongwe (Kamuzu International Airport)", "FWCL": "Blantyre (Chileka Airport)"}


def station_label(icao):
    return STATION_NAMES.get(icao.upper(), icao.upper())


# %% Criteria (ICAO Annex 3) - EDIT THRESHOLDS HERE to suit local conditions
CRITERIA = dict(
    # visibility below x metres. (threshold, level) - lowest threshold = most severe
    vis_bands_m=[(1000, "very_extreme"), (3000, "extreme"), (5000, "significant")],
    # cloud ceiling (BKN/OVC/VV) at or below x feet
    ceiling_bands_ft=[(100, "extreme"), (200, "extreme"), (500, "significant"),
                      (1000, "significant"), (1500, "significant")],
    dir_change_deg=60, dir_change_min_kt=10,     # wind direction change between consecutive reports
    speed_change_kt=10,                          # mean wind speed change between consecutive reports
    gust_increase_kt=10, gust_min_mean_kt=15,    # gust spread increase, with mean speed at least this
)

# event tag -> (level, plain-language name, technical definition). Thresholds appear in the bands above.
FIXED_EVENTS = {
    "THUNDERSTORM": ("extreme", "Thunderstorms", "Thunderstorm at the aerodrome (TS in present weather)."),
    "HEAVY_PRECIP": ("extreme", "Heavy rain", "Heavy rain or drizzle (+RA, +TSRA, +SHRA, +DZ)."),
    "MOD_PRECIP": ("significant", "Moderate rain", "Moderate precipitation (RA/DZ etc. without an intensity sign)."),
    "HAIL": ("extreme", "Hail", "Hail or small hail (GR/GS)."),
    "FUNNEL_CLOUD": ("extreme", "Funnel cloud / tornado", "Funnel cloud, tornado or waterspout (FC)."),
    "SQUALL": ("extreme", "Squalls", "Squall (SQ)."),
    "DUST_SAND_STORM": ("extreme", "Dust / sand storm", "Duststorm or sandstorm (DS/SS)."),
    "FREEZING_PRECIP": ("extreme", "Freezing rain", "Freezing rain or drizzle (FZRA/FZDZ)."),
    "FOG": ("significant", "Fog", "Fog (FG, excluding shallow MIFG)."),
    "WIND_SHIFT": ("significant", "Sudden wind direction change", ""),
    "WIND_SPEED_CHANGE": ("significant", "Sudden change in wind speed", ""),
    "GUST_INCREASE": ("significant", "Gusty winds", ""),
}
GROUPS = ["Storms and rain", "Wind", "Low visibility / fog", "Low cloud"]


def vis_tag(thr):
    return f"VIS_LT_{int(thr)}"


def cig_tag(thr):
    return f"CIG_LE_{int(thr)}FT"


def group_of(tag):
    if tag.startswith("VIS_") or tag in ("FOG", "DUST_SAND_STORM"):
        return "Low visibility / fog"
    if tag.startswith("CIG_"):
        return "Low cloud"
    if tag in ("WIND_SHIFT", "WIND_SPEED_CHANGE", "GUST_INCREASE", "SQUALL"):
        return "Wind"
    return "Storms and rain"


def event_family(tag, names):
    """Plain-language family: all visibility bands -> 'Low visibility', all ceiling bands -> 'Low cloud'."""
    if tag.startswith("VIS_"):
        return "Low visibility"
    if tag.startswith("CIG_"):
        return "Low cloud"
    return names[tag]


def event_catalog(C):
    """Return (levels, plain names, definitions) for every event tag, using the thresholds in C."""
    levels = {k: v[0] for k, v in FIXED_EVENTS.items()}
    names = {k: v[1] for k, v in FIXED_EVENTS.items()}
    info = {k: v[2] for k, v in FIXED_EVENTS.items()}
    for thr, lv in C["vis_bands_m"]:
        t = vis_tag(thr)
        levels[t], names[t], info[t] = lv, f"Visibility below {thr / 1000:g} km", f"Prevailing visibility below {thr:g} m."
    for thr, lv in C["ceiling_bands_ft"]:
        t = cig_tag(thr)
        levels[t], names[t], info[t] = lv, f"Low cloud (ceiling at or below {thr:g} ft)", \
            f"Cloud ceiling (BKN/OVC/VV) at or below {thr:g} ft."
    info["WIND_SHIFT"] = (f"Mean wind direction changed by {C['dir_change_deg']} degrees or more between consecutive "
                          f"reports (mean speed {C['dir_change_min_kt']} kt or more before and/or after).")
    info["WIND_SPEED_CHANGE"] = f"Mean wind speed changed by {C['speed_change_kt']} kt or more between consecutive reports."
    info["GUST_INCREASE"] = (f"Gust spread (gust minus mean) rose by {C['gust_increase_kt']} kt or more between "
                             f"consecutive reports, with mean speed of {C['gust_min_mean_kt']} kt or more.")
    return levels, names, info


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
CLD_RE = re.compile(r"^(FEW|SCT|BKN|OVC|VV)(\d{3})(CB|TCU)?$")
VIS_RE = re.compile(r"^\d{4}$")
WX_RE = re.compile(r"^(\+|-|VC)?(MI|BC|PR|DR|BL|SH|TS|FZ)*"
                   r"(DZ|RA|SN|SG|PL|GR|GS|UP|BR|FG|FU|VA|DU|SA|HZ|PO|SQ|FC|SS|DS)*$")
STOP = {"RMK", "TEMPO", "BECMG", "NOSIG"}
NUM_COLS = ["wind_dir", "wind_kt", "gust_kt", "vis_m", "temp_c", "dew_c", "ceiling_ft"]


def _temp(s):
    return int(s.replace("M", "-")) if s else None


def parse_report(stamp, rtype, body):
    rec = dict(time=pd.to_datetime(stamp, format="%Y%m%d%H%M"), rtype=rtype,
               wind_dir=None, wind_kt=None, gust_kt=None, vis_m=None, temp_c=None, dew_c=None,
               ceiling_ft=None, wx="", raw=body)
    wx = []
    for t in body.split():
        if t in STOP:
            break
        if t == "CAVOK":
            rec["vis_m"] = 9999
        elif (m := WIND_RE.match(t)):
            k = MPS_TO_KT if m.group(4) == "MPS" else 1
            rec["wind_dir"] = None if m.group(1) == "VRB" else int(m.group(1))
            rec["wind_kt"] = round(int(m.group(2)) * k)
            rec["gust_kt"] = round(int(m.group(3)) * k) if m.group(3) else None
        elif (m := TEMP_RE.match(t)):
            rec["temp_c"], rec["dew_c"] = _temp(m.group(1)), _temp(m.group(2))
        elif (m := CLD_RE.match(t)):
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


# %% Detect extreme events
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
    if etype.startswith("VIS_"):
        return "vis_m", "min"
    if etype.startswith("CIG_"):
        return "ceiling_ft", "min"
    return None


def band_tag(value, bands, tagger, inclusive):
    for thr, _lv in sorted(bands):
        if (value <= thr) if inclusive else (value < thr):
            return tagger(thr)
    return None


def tag_report(r, C):
    tags, wx = [], r["wx"].split()
    if pd.notna(r["vis_m"]) and (t := band_tag(r["vis_m"], C["vis_bands_m"], vis_tag, False)):
        tags.append(t)
    if pd.notna(r["ceiling_ft"]) and (t := band_tag(r["ceiling_ft"], C["ceiling_bands_ft"], cig_tag, True)):
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


# %% Run the analysis and print a summary
def analyse(icao, year, month, gap_hours=3, criteria=None):
    """criteria: optional dict overriding entries of CRITERIA for this run only."""
    icao = icao.upper()
    C = {**CRITERIA, **(criteria or {})}
    text = fetch_metars(icao, year, month)
    try:
        df = parse_metars(text)
    except SystemExit:
        (Path("cache") / f"{icao}_{year}{month:02d}.txt").unlink(missing_ok=True)   # drop bad download
        raise
    levels, names, info = event_catalog(C)
    tags = [tag_report(r, C) for r in df.to_dict("records")]
    shift, spd, gust = change_tags(df, C["dir_change_deg"], C["dir_change_min_kt"], C["speed_change_kt"],
                                   C["gust_increase_kt"], C["gust_min_mean_kt"])
    for i in range(len(df)):
        if shift.iloc[i]:
            tags[i].append("WIND_SHIFT")
        if spd.iloc[i]:
            tags[i].append("WIND_SPEED_CHANGE")
        if gust.iloc[i]:
            tags[i].append("GUST_INCREASE")
    ev = explode_events(df, tags, levels)
    ep = build_episodes(ev, levels, gap_hours)
    return dict(icao=icao, year=year, month=month, df=df, ev=ev, ep=ep, criteria=C,
                tag_level=levels, names=names, info=info)


KEY_PARAMS = (("Highest temperature", "temp_c", "idxmax", "°C"), ("Lowest temperature", "temp_c", "idxmin", "°C"),
              ("Strongest wind / gust", "wind_peak_kt", "idxmax", "kt"),
              ("Lowest visibility", "vis_m", "idxmin", "m"), ("Lowest cloud ceiling", "ceiling_ft", "idxmin", "ft"))


def summarise(res):
    df, ep, icao, year, month = res["df"], res["ep"], res["icao"], res["year"], res["month"]
    days = calendar.monthrange(year, month)[1]
    print(f"\n=== {icao}  {year}-{month:02d} ===")
    print(f"Reports: {len(df)} (≈{len(df) / (days * 24) * 100:.0f}% of hourly expected), "
          f"of which SPECI: {(df['rtype'] == 'SPECI').sum()}")
    for label, col, fn, unit in KEY_PARAMS:
        if df[col].notna().any():
            i = getattr(df[col], fn)()
            print(f"{label}: {df.at[i, col]:.0f} {unit} at {df.at[i, 'time']}")
    print("\nEpisodes by event type:")
    if ep.empty:
        print("  none detected")
    else:
        print(ep.assign(level=ep["level"].map(LEVEL_LABELS), event=ep["event"].map(res["names"]))
              .groupby(["level", "event"]).size().rename("episodes").to_string())


# %% Graphs
def _finish(fig, out):
    import matplotlib.pyplot as plt
    fig.tight_layout()
    if out:
        fig.savefig(out, dpi=150)
        plt.close(fig)
        print(f"Plot saved: {out}")
    else:
        plt.show()


def make_plot(res, out=None):
    """Temperature, wind/gust, visibility and ceiling with event periods shaded."""
    import matplotlib.pyplot as plt
    df, ep = res["df"], res["ep"]
    fig, ax = plt.subplots(4, 1, figsize=(11, 10), sharex=True)
    ax[0].plot(df["time"], df["temp_c"], lw=1, color="tab:red"); ax[0].set_ylabel("Temp (°C)")
    ax[1].plot(df["time"], df["wind_peak_kt"], lw=1, color="tab:blue"); ax[1].set_ylabel("Wind/gust (kt)")
    ax[2].plot(df["time"], df["vis_m"], lw=1, color="tab:gray"); ax[2].set_ylabel("Visibility (m)")
    ax[3].plot(df["time"], df["ceiling_ft"], ".", ms=3, color="tab:green"); ax[3].set_ylabel("Ceiling (ft)")
    panel = {"Wind": 1, "Low visibility / fog": 2, "Low cloud": 3}
    for _, e in ep.iterrows():
        p = panel.get(group_of(e["event"]), 1)
        ax[p].axvspan(e["start"], e["end"] + pd.Timedelta(hours=1), alpha=0.35, color=LEVEL_COLOR[e["level"]])
    fig.suptitle(f"{res['icao']} {res['year']}-{res['month']:02d} - observations with extreme events shaded\n"
                 "(dark red = very extreme, red = extreme, orange = significant)")
    _finish(fig, out)


def plot_timeline(res, out=None):
    """One row per event type, one tick per flagged report."""
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt
    ev, year, month = res["ev"], res["year"], res["month"]
    if ev.empty:
        print("No events to plot.")
        return
    order = (ev.assign(s=ev["level"].map(LEVEL_SCORE)).groupby("event")["s"].max().sort_values().index.tolist())
    ypos = {e: i for i, e in enumerate(order)}
    fig, ax = plt.subplots(figsize=(12, 0.55 * len(order) + 2.2))
    for level in ("significant", "extreme", "very_extreme"):
        g = ev[ev["level"] == level]
        if not g.empty:
            ax.scatter(g["time"], g["event"].map(ypos), marker="|", s=450, linewidths=2.2,
                       color=LEVEL_COLOR[level], label=LEVEL_LABELS[level])
    start = pd.Timestamp(year, month, 1)
    ax.set_xlim(start, start + pd.offsets.MonthEnd(0) + pd.Timedelta(days=1))
    ax.set_yticks(range(len(order))); ax.set_yticklabels([res["names"][e] for e in order])
    ax.set_ylim(-0.7, len(order) - 0.3)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d"))
    ax.tick_params(axis="x", labelsize=7)
    ax.grid(axis="x", alpha=0.3)
    ax.set_xlabel(f"Day of month, {year}-{month:02d} (UTC)")
    ax.set_title(f"{res['icao']} - when each event occurred")
    ax.legend(loc="upper right", ncol=3, fontsize=8, framealpha=0.9)
    _finish(fig, out)


def _grid_plot(res, mode, figsize, title=""):
    """mode 'hour': rows = hours (day x hour map). mode 'group': rows = event groups (bulletin calendar)."""
    import numpy as np
    import matplotlib.pyplot as plt
    from matplotlib.colors import ListedColormap
    from matplotlib.patches import Patch
    df, ev, year, month = res["df"], res["ev"], res["year"], res["month"]
    days = calendar.monthrange(year, month)[1]
    rows = 24 if mode == "hour" else len(GROUPS)
    row_of = (lambda t, e: t.hour) if mode == "hour" else (lambda t, e: GROUPS.index(group_of(e)))
    grid = np.zeros((rows, days))
    if mode == "hour":
        for t in df["time"]:
            grid[t.hour, t.day - 1] = 1
    else:
        for d in {t.day for t in df["time"]}:
            grid[:, d - 1] = 1
    for _, r in ev.iterrows():
        i, d = row_of(r["time"], r["event"]), r["time"].day - 1
        grid[i, d] = max(grid[i, d], 1 + LEVEL_SCORE[r["level"]])
    colors = ["white", "#e6e6e6", "orange", "red", "darkred"]
    fig, ax = plt.subplots(figsize=figsize)
    ax.imshow(grid, aspect="auto", origin="lower" if mode == "hour" else "upper",
              cmap=ListedColormap(colors), vmin=0, vmax=4, interpolation="nearest")
    ax.set_xticks(range(days)); ax.set_xticklabels(range(1, days + 1), fontsize=7)
    if mode == "hour":
        ax.set_yticks(range(0, 24, 2)); ax.set_yticklabels([f"{h:02d}" for h in range(0, 24, 2)], fontsize=8)
        ax.set_ylabel("Hour (UTC)")
    else:
        ax.set_yticks(range(rows)); ax.set_yticklabels(GROUPS, fontsize=8)
    ax.set_xticks(np.arange(-0.5, days, 1), minor=True); ax.set_yticks(np.arange(-0.5, rows, 1), minor=True)
    ax.grid(which="minor", color="white", lw=0.8); ax.tick_params(which="minor", length=0)
    ax.set_xlabel(f"Day of month, {year}-{month:02d}")
    if title:
        ax.set_title(title)
    names = ["no report / NIL", "report, no extreme", "significant", "extreme", "very extreme"]
    ax.legend(handles=[Patch(facecolor=c, edgecolor="grey", label=n) for c, n in zip(colors, names)],
              loc="upper center", bbox_to_anchor=(0.5, -0.12 if mode == "hour" else -0.28),
              ncol=5, fontsize=8, frameon=False)
    return fig


def plot_heatmap(res, out=None):
    """Day x hour grid: worst event in each hour."""
    _finish(_grid_plot(res, "hour", (12, 6.5), f"{res['icao']} - worst event by day and hour"), out)


def plot_calendar(res, out=None):
    """Compact day x event-group calendar used in the bulletin."""
    _finish(_grid_plot(res, "group", (8.4, 2.9)), out)


# %% Word bulletin (2 pages)
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


def write_bulletin(res, outdir="output"):
    """Two-page public bulletin (Word). Returns the file path."""
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    icao, year, month, df, ev, ep = res["icao"], res["year"], res["month"], res["df"], res["ev"], res["ep"]
    names, tl, C = res["names"], res["tag_level"], res["criteria"]
    outdir = Path(outdir); outdir.mkdir(exist_ok=True)
    stem = f"{icao}_{year}{month:02d}"
    mname = f"{calendar.month_name[month]} {year}"
    NAVY = "1F3A5F"
    fmt = lambda t: f"{t.day} {t:%b} {t:%H:%M}Z"

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.5)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(9.5)
    st.paragraph_format.space_after = Pt(3)
    for name, size in (("Heading 1", 12.5), ("Heading 2", 11)):
        h = doc.styles[name]
        h.font.name, h.font.size, h.font.bold = "Calibri", Pt(size), True
        h.font.color.rgb = RGBColor.from_string(NAVY)
        h.paragraph_format.space_before, h.paragraph_format.space_after = Pt(8), Pt(3)

    logo = find_logo()
    if logo:
        sec.top_margin, sec.header_distance = Cm(2.6), Cm(0.7)
        sec.header.paragraphs[0].add_run().add_picture(str(logo), width=Cm(5.2))
    else:
        print(f"[note] {LOGO_FILE} not found - bulletin written without the DCCMS logo.")

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.add_run(f"Extreme weather bulletin - {icao} - {mname}   |   Page ").font.size = Pt(8)
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

    def lines(cell, items):
        """items: list of (text, size, bold, hex colour, align)"""
        cell.text = ""
        for i, (text, size, bold, color, align) in enumerate(items):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            if align:
                p.alignment = align
            r = p.add_run(str(text)); r.font.size, r.bold = Pt(size), bold
            if color:
                r.font.color.rgb = RGBColor.from_string(color)

    def table(header, rows, widths, level_col=None, size=8.5):
        t = doc.add_table(rows=1, cols=len(header))
        t.style, t.alignment, t.autofit = "Table Grid", WD_TABLE_ALIGNMENT.CENTER, False
        for i, h in enumerate(header):
            lines(t.rows[0].cells[i], [(h, size, True, "FFFFFF", None)]); shade(t.rows[0].cells[i], NAVY)
        for row in rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                lines(cells[i], [(v, size, False, None, None)])
                if level_col is not None and i == level_col:
                    lv = next((k for k, lab in LEVEL_LABELS.items() if lab == v), None)
                    if lv:
                        shade(cells[i], LEVEL_FILL[lv])
                        lines(cells[i], [(v, size, True, "FFFFFF" if lv == "very_extreme" else None, None)])
        for i, w in enumerate(widths):
            t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
        return t

    # ---------- derived numbers ----------
    ev = ev.copy()
    ev["score"] = ev["level"].map(LEVEL_SCORE)
    first_t, last_t = df["time"].min(), df["time"].max()
    exp = calendar.monthrange(year, month)[1] * 24
    cover = len(df) / exp * 100
    ep = ep.copy()
    if len(ep):
        ep["score"] = ep["level"].map(LEVEL_SCORE)
        ep["dur"] = ep["end"] - ep["start"]
    n_by = ep.groupby("level").size().to_dict() if len(ep) else {}

    # ---------- banner ----------
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment, banner.autofit = WD_TABLE_ALIGNMENT.CENTER, False
    cell = banner.rows[0].cells[0]
    shade(cell, NAVY)
    lines(cell, [("MONTHLY EXTREME WEATHER BULLETIN", 20, True, "FFFFFF", None),
                 (f"{station_label(icao)}  |  {mname}", 12, False, "DCE6F2", None)])
    banner.columns[0].width = cell.width = Cm(17.4)
    p = doc.add_paragraph()
    r = p.add_run(f"Observations {first_t.day} {first_t:%b} - {last_t.day} {last_t:%b %Y}  |  Source: aerodrome weather "
                  f"reports (METAR/SPECI)  |  Times in UTC (Malawi local time = UTC+2)  |  Issued {pd.Timestamp.now():%d %b %Y}")
    r.italic, r.font.size = True, Pt(8)

    # ---------- at a glance ----------
    doc.add_heading("Weather at a glance", level=1)
    tiles = []
    for label, col, fn, unit in KEY_PARAMS:
        if df[col].notna().any():
            i = getattr(df[col], fn)()
            tiles.append((label, f"{df.at[i, col]:.0f} {unit}", fmt(df.at[i, "time"])))
        else:
            tiles.append((label, "n/a", ""))
    tiles.append(("Extreme-weather episodes", str(len(ep)),
                  f"{n_by.get('very_extreme', 0)} very extreme, {n_by.get('extreme', 0)} extreme, "
                  f"{n_by.get('significant', 0)} significant"))
    t = doc.add_table(rows=1, cols=len(tiles))
    t.alignment, t.autofit = WD_TABLE_ALIGNMENT.CENTER, False
    for i, (label, value, sub) in enumerate(tiles):
        c = t.rows[0].cells[i]
        shade(c, "EAF1FA")
        lines(c, [(label, 7.5, False, "595959", WD_ALIGN_PARAGRAPH.CENTER),
                  (value, 14, True, NAVY, WD_ALIGN_PARAGRAPH.CENTER),
                  (sub, 7, False, "595959", WD_ALIGN_PARAGRAPH.CENTER)])
        t.columns[i].width = c.width = Cm(17.4 / len(tiles))

    # ---------- summary ----------
    doc.add_heading("Summary", level=1)
    if ep.empty:
        doc.add_paragraph(f"No significant or extreme weather events were recorded at {station_label(icao)} "
                          f"during {mname}, using the thresholds listed on page 2.")
    else:
        fam_n = ep.assign(fam=ep["event"].map(lambda e: event_family(e, names))).groupby("fam").size()
        top = [f"{f.lower()} ({n})" for f, n in fam_n.sort_values(ascending=False).head(3).items()]
        dd = ev.assign(day=ev["time"].dt.date).groupby("day").agg(s=("score", "max"), n=("event", "size"))
        wd = dd.sort_values(["s", "n"], ascending=False).index[0]
        day_ev = ev[ev["time"].dt.date == wd].sort_values("score", ascending=False)
        types_wd = day_ev["event"].map(lambda e: event_family(e, names)).drop_duplicates()
        worst = ep.sort_values(["score", "dur"], ascending=False).iloc[0]
        spec = peak_spec(worst["event"])
        if spec and pd.notna(worst.get("peak")):
            detail = {"vis_m": f"visibility fell to {worst['peak']:.0f} m", "ceiling_ft": f"cloud ceiling fell to {worst['peak']:.0f} ft"}[spec[0]]
        elif worst["reports"] > 1:
            detail = f"lasting about {max(1, round(worst['dur'].total_seconds() / 3600))} hours"
        else:
            detail = "reported once"
        doc.add_paragraph(
            f"During {mname}, {station_label(icao)} recorded {len(ep)} episodes of significant or extreme weather: "
            f"{n_by.get('very_extreme', 0)} very extreme, {n_by.get('extreme', 0)} extreme and "
            f"{n_by.get('significant', 0)} significant. The most frequent were {', '.join(top)}.")
        doc.add_paragraph(
            f"The most affected day was {wd.day} {wd:%B}, with {', '.join(f.lower() for f in types_wd.head(3))}"
            f"{' and more' if len(types_wd) > 3 else ''}. The most severe episode was "
            f"{names[worst['event']].lower()} on {worst['start'].day} {worst['start']:%B} ({detail}).")

    # ---------- page 2 ----------
    doc.add_heading("When extremes occurred", level=1)
    cal = outdir / f"{stem}_calendar.png"
    plot_calendar(res, cal)
    doc.add_picture(str(cal), width=Cm(17.0))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()
    # ---------- events table ----------
    doc.add_heading("Events recorded", level=1)
    if ev.empty:
        doc.add_paragraph("No events.")
    else:
        stats = []
        for etype, g in ev.groupby("event"):
            days = sorted(g["time"].dt.day.unique())
            shown = ", ".join(str(d) for d in days[:14]) + (f" (+{len(days) - 14} more)" if len(days) > 14 else "")
            stats.append((LEVEL_SCORE[tl[etype]], int((ep["event"] == etype).sum()), names[etype], LEVEL_LABELS[tl[etype]], shown))
        stats.sort(key=lambda s: (-s[0], -s[1], s[2]))
        cap = 10
        rows = [(n, lab, eps, d) for _, eps, n, lab, d in stats[:cap]]
        table(["Event", "Level", "Episodes", f"Days of {mname.split()[0]} affected"], rows, [5.6, 2.3, 1.8, 7.7], level_col=1)
        if len(stats) > cap:
            doc.add_paragraph(f"... and {len(stats) - cap} other, less severe event types.").runs[0].italic = True

    doc.add_heading("Most severe episodes", level=1)
    if ep.empty:
        doc.add_paragraph("No episodes.")
    else:
        rows = []
        for _, e in ep.sort_values(["score", "dur"], ascending=False).head(6).iterrows():
            when = fmt(e["start"]) if e["start"] == e["end"] else (
                f"{fmt(e['start'])} to {e['end']:%H:%M}Z" if e["start"].date() == e["end"].date()
                else f"{fmt(e['start'])} to {fmt(e['end'])}")
            spec = peak_spec(e["event"])
            det = "-"
            if spec and pd.notna(e.get("peak")):
                det = f"lowest {e['peak']:.0f} {'m' if spec[0] == 'vis_m' else 'ft'}"
            dur = "single report" if e["reports"] == 1 else f"{max(1, round(e['dur'].total_seconds() / 3600))} h"
            rows.append((names[e["event"]], LEVEL_LABELS[e["level"]], when, dur, det))
        table(["Event", "Level", "When (UTC)", "Duration", "Detail"], rows, [5.0, 2.2, 5.4, 2.2, 2.6], level_col=1)

    doc.add_heading("Most affected days", level=1)
    if ev.empty:
        doc.add_paragraph("No flagged days.")
    else:
        rows = []
        dsum = ev.assign(day=ev["time"].dt.date).groupby("day").agg(s=("score", "max"), n=("event", "size"))
        for day in dsum.sort_values(["s", "n"], ascending=False).head(5).index:
            g = ev[ev["time"].dt.date == day].drop_duplicates("event").sort_values("score", ascending=False)
            nm = list(g["event"].map(lambda e: event_family(e, names)).drop_duplicates())
            txt = ", ".join(nm[:3]) + (f" (+{len(nm) - 3} more)" if len(nm) > 3 else "")
            worst_lv = {v: k for k, v in LEVEL_SCORE.items()}[int(dsum.loc[day, "s"])]
            rows.append((f"{day:%a} {day.day} {day:%b}", LEVEL_LABELS[worst_lv], txt))
        table(["Date", "Worst level", "Events"], rows, [3.0, 2.6, 11.8], level_col=1)

    doc.add_heading("About this bulletin", level=2)
    vb = ", ".join(f"{t / 1000:g}" for t, _ in sorted(C["vis_bands_m"]))
    cb = ", ".join(f"{t:g}" for t, _ in sorted(C["ceiling_bands_ft"]))
    for txt in (
        f"Events are identified from routine aerodrome weather reports using thresholds based on ICAO Annex 3 "
        f"(SPECI criteria), adjusted for local conditions: visibility below {vb} km; cloud ceiling at or below {cb} ft; "
        f"thunderstorm, heavy rain, hail, squall, funnel cloud, dust or sand storm, freezing rain and fog; sudden wind "
        f"changes (direction {C['dir_change_deg']}° or more, speed {C['speed_change_kt']} kt or more, increase in gusts "
        f"of {C['gust_increase_kt']} kt or more). Events within 3 hours of each other count as one episode.",
        f"The report covers about {cover:.0f}% of the hourly observations expected for the month. It describes conditions "
        f"at the aerodrome, which can differ from nearby areas. For forecasts and warnings, refer to official DCCMS products."):
        p = doc.add_paragraph(txt)
        for r in p.runs:
            r.font.size = Pt(8)

    path = outdir / f"{stem}_bulletin.docx"
    doc.save(path)
    return path


# %% Save CSVs, graphs and bulletin
def save_outputs(res, outdir="output", plots=True, report=True):
    out = Path(outdir); out.mkdir(exist_ok=True)
    stem = f"{res['icao']}_{res['year']}{res['month']:02d}"
    res["df"].to_csv(out / f"{stem}_parsed.csv", index=False)
    res["ev"].to_csv(out / f"{stem}_flagged_reports.csv", index=False)
    res["ep"].to_csv(out / f"{stem}_episodes.csv", index=False)
    if plots or report:
        make_plot(res, out / f"{stem}_values.png")
        plot_timeline(res, out / f"{stem}_timeline.png")
        plot_heatmap(res, out / f"{stem}_heatmap.png")
    if report:
        path = write_bulletin(res, out)
        print("Bulletin saved:", path)
        return path
    print("Saved to", out.resolve())


# %% CLI
def main(argv=None):
    import argparse
    p = argparse.ArgumentParser(description="Ogimet METAR extreme-event analysis (ICAO Annex 3 criteria)")
    p.add_argument("icao"); p.add_argument("year", type=int); p.add_argument("month", type=int)
    p.add_argument("--gap", type=float, default=3, help="hours gap that splits episodes")
    p.add_argument("--outdir", default="output")
    p.add_argument("--plot", action="store_true", help="save graphs")
    p.add_argument("--report", action="store_true", help="write the 2-page Word bulletin (needs python-docx)")
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
