"""
METAR Extreme Weather Analyser - web app (Streamlit).

Run locally:   streamlit run app.py
Needs metar_extremes.py and dccms_logo.png in the same folder.
"""
import datetime as dt
import io
import re
import tempfile
import threading
import zipfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import pandas as pd
import streamlit as st

import metar_extremes as me

st.set_page_config(page_title="METAR Extreme Weather Analyser", page_icon="⛈️", layout="wide")

STATIONS = {f"{code} - {name}": code for code, name in me.STATION_NAMES.items()}
STATIONS["Other airport (type the ICAO code)"] = None
today = dt.date.today()
prev = today.replace(day=1) - dt.timedelta(days=1)          # default: last completed month


@st.cache_resource
def get_lock():
    return threading.Lock()                                  # one analysis at a time (matplotlib is not thread-safe)


def key_observations(df):
    rows = []
    for label, col, fn, unit in me.KEY_PARAMS:
        if df[col].notna().any():
            i = getattr(df[col], fn)()
            rows.append({"Parameter": label, "Value": f"{df.at[i, col]:.0f} {unit}",
                         "Time (UTC)": df.at[i, "time"].strftime("%d %b %H:%M") + "Z"})
    return pd.DataFrame(rows)


def run_job(icao, year, month, gap, criteria):
    """Analyse, then write graphs + Word bulletin + CSVs into a private temp folder."""
    with get_lock():
        res = me.analyse(icao, year, month, gap, criteria=criteria)
        outdir = Path(tempfile.mkdtemp(prefix="metar_"))
        report_path = me.save_outputs(res, outdir, plots=True, report=True)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(outdir.iterdir()):
                z.write(f, f.name)
    return dict(res=res, outdir=outdir, report_name=report_path.name,
                report=report_path.read_bytes(), zip=buf.getvalue())


def band_inputs(label, bands, unit, min_value):
    """Number inputs for a list of (threshold, level) bands. Levels keep their order; thresholds are sorted."""
    vals = []
    for thr, lv in bands:
        vals.append(st.number_input(f"{label} - {me.LEVEL_LABELS[lv].lower()} ({unit})", value=thr, min_value=min_value))
    return list(zip(sorted(vals), [lv for _, lv in bands]))


# ------------------------------------------------------------------ sidebar
LOGO = me.find_logo()
if LOGO:
    st.image(str(LOGO), width=360)
st.title("⛈️ METAR Extreme Weather Analyser")
st.caption("Downloads a month of METAR/SPECI reports from Ogimet, finds extreme weather events using ICAO Annex 3 "
           "criteria (adjustable for local conditions) and writes a 2-page bulletin. All times are UTC.")

sb = st.sidebar
sb.header("1. Choose station and month")
choice = sb.selectbox("Station", list(STATIONS))
icao = STATIONS[choice]
if icao is None:
    icao = sb.text_input("ICAO code (4 letters)", value="FWKI").strip().upper()
years = list(range(today.year, 1999, -1))
year = sb.selectbox("Year", years, index=years.index(prev.year))
month = sb.selectbox("Month", list(range(1, 13)), index=prev.month - 1,
                     format_func=lambda m: pd.Timestamp(2000, m, 1).strftime("%B"))

sb.header("2. Thresholds (optional)")
C = dict(me.CRITERIA)
with sb.expander("ICAO thresholds - adjust for local conditions"):
    st.caption("Defaults follow ICAO Annex 3. Change a value to suit local conditions.")
    C["vis_bands_m"] = band_inputs("Visibility below", me.CRITERIA["vis_bands_m"], "m", 50)
    C["ceiling_bands_ft"] = band_inputs("Ceiling at or below", me.CRITERIA["ceiling_bands_ft"], "ft", 10)
    C["dir_change_deg"] = st.number_input("Wind direction change (degrees)", value=C["dir_change_deg"], min_value=10)
    C["dir_change_min_kt"] = st.number_input("...with mean speed of at least (kt)", value=C["dir_change_min_kt"], min_value=1)
    C["speed_change_kt"] = st.number_input("Wind speed change (kt)", value=C["speed_change_kt"], min_value=1)
    C["gust_increase_kt"] = st.number_input("Gust increase (kt)", value=C["gust_increase_kt"], min_value=1)
    C["gust_min_mean_kt"] = st.number_input("...with mean speed of at least (kt)", value=C["gust_min_mean_kt"], min_value=1)
gap = sb.number_input("Hours gap that splits one event into two", value=3, min_value=1, max_value=24)

go = sb.button("Analyse", type="primary")

# ------------------------------------------------------------------ run
if go:
    if not re.fullmatch(r"[A-Z]{4}", icao or ""):
        st.error("Enter a valid 4-letter ICAO code, for example FWKI.")
    elif (year, month) > (today.year, today.month):
        st.error("That month is in the future. Choose a past or current month.")
    else:
        if (year, month) == (today.year, today.month):
            st.info("This month is not finished yet, so the results cover only the days so far.")
        try:
            with st.spinner(f"Downloading and analysing {icao} {year}-{month:02d} ... this can take a minute"):
                st.session_state["job"] = run_job(icao, year, month, float(gap), C)
        except SystemExit as e:
            st.session_state.pop("job", None)
            st.error(f"{e}  Ogimet may be busy or have no data for this station and month. Wait a few minutes and try again.")
        except Exception as e:
            st.session_state.pop("job", None)
            st.error(f"Something went wrong: {e}")

job = st.session_state.get("job")
if not job:
    st.info("Choose a station and month in the sidebar, then press **Analyse**.")
    st.stop()

res, outdir = job["res"], job["outdir"]
df, ep, ev, names = res["df"], res["ep"], res["ev"], res["names"]
stem = f"{res['icao']}_{res['year']}{res['month']:02d}"
days = pd.Period(f"{res['year']}-{res['month']:02d}").days_in_month

st.subheader(f"{me.station_label(res['icao'])} - {pd.Timestamp(res['year'], res['month'], 1).strftime('%B %Y')}")
d1, d2 = st.columns(2)
d1.download_button("⬇️ Download 2-page bulletin (Word)", job["report"], file_name=job["report_name"],
                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
d2.download_button("⬇️ Download everything (ZIP: bulletin, graphs, CSVs)", job["zip"],
                   file_name=f"{stem}_results.zip", mime="application/zip")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Reports analysed", len(df))
m2.metric("Coverage of hourly reports", f"{len(df) / (days * 24) * 100:.0f}%")
m3.metric("Extreme-weather episodes", len(ep))
m4.metric("Very extreme episodes", int((ep["level"] == "very_extreme").sum()) if len(ep) else 0)

tabs = st.tabs(["Overview", "Events", "Graphs"])

with tabs[0]:
    st.markdown("**Key observations**")
    st.dataframe(key_observations(df))

with tabs[1]:
    if ep.empty:
        st.success("No events detected with these thresholds.")
    else:
        e2 = ep.assign(level=ep["level"].map(me.LEVEL_LABELS), event=ep["event"].map(names))
        st.markdown("**Summary by event type**")
        st.dataframe(e2.groupby(["level", "event"]).agg(episodes=("start", "size"), reports=("reports", "sum"),
                                                        first=("start", "min"), last=("end", "max")).reset_index())
        st.markdown("**All episodes**")
        st.dataframe(e2)
        with st.expander("Raw METAR for every flagged report"):
            raw = ev[["time", "rtype", "event", "level", "raw"]]
            st.dataframe(raw.assign(level=raw["level"].map(me.LEVEL_LABELS), event=raw["event"].map(names)))
        with st.expander("Definitions used"):
            for tag in sorted(res["info"], key=lambda t: (-me.LEVEL_SCORE[res["tag_level"][t]], t)):
                st.write(f"**{names[tag]}** ({me.LEVEL_LABELS[res['tag_level'][tag]]}) - {res['info'][tag]}")

with tabs[2]:
    for suffix, cap in (("calendar", "When extremes occurred (bulletin chart)"),
                        ("timeline", "Timeline of every flagged report"),
                        ("heatmap", "Worst event by day and hour"),
                        ("values", "Observations with extreme events shaded")):
        f = outdir / f"{stem}_{suffix}.png"
        if f.exists():
            st.image(str(f), caption=cap)

st.caption("Thresholds follow ICAO Annex 3 and can be adjusted for local conditions. "
           "Results depend on report availability; NIL reports are excluded.")
