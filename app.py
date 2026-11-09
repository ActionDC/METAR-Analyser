"""
METAR Extreme Weather Analyser - web app (Streamlit).

Run locally:   streamlit run app.py
Needs metar_extremes.py in the same folder.
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

STATIONS = {
    "FWKI - Lilongwe (Kamuzu Intl)": "FWKI",
    "FWCL - Blantyre (Chileka)": "FWCL",
    "Other airport (type the ICAO code)": None,
}
MONTHS = list(range(1, 13))
today = dt.date.today()
prev = today.replace(day=1) - dt.timedelta(days=1)          # default: last completed month


@st.cache_resource
def get_lock():
    return threading.Lock()                                  # one analysis at a time (matplotlib is not thread-safe)


def key_observations(df):
    rows = []
    for label, col, fn, unit in (("Highest temperature", "temp_c", "idxmax", "°C"),
                                 ("Lowest temperature", "temp_c", "idxmin", "°C"),
                                 ("Strongest wind / gust", "wind_peak_kt", "idxmax", "kt"),
                                 ("Lowest visibility", "vis_m", "idxmin", "m"),
                                 ("Lowest cloud ceiling", "ceiling_ft", "idxmin", "ft"),
                                 ("Lowest QNH", "qnh_hpa", "idxmin", "hPa")):
        if df[col].notna().any():
            i = getattr(df[col], fn)()
            rows.append({"Parameter": label, "Value": f"{df.at[i, col]:.0f} {unit}",
                         "Time (UTC)": df.at[i, "time"].strftime("%d %b %H:%M") + "Z"})
    return pd.DataFrame(rows)


def run_job(icao, year, month, gap, local):
    """Analyse, then write graphs + Word report + CSVs into a private temp folder."""
    with get_lock():
        res = me.analyse(icao, year, month, gap, local=local)
        outdir = Path(tempfile.mkdtemp(prefix="metar_"))
        report_path = me.save_outputs(res, outdir, plots=True, report=True)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(outdir.iterdir()):
                z.write(f, f.name)
    return dict(res=res, outdir=outdir, report_name=report_path.name,
                report=report_path.read_bytes(), zip=buf.getvalue())


# ------------------------------------------------------------------ sidebar
LOGO = me.find_logo()
if LOGO:
    st.image(str(LOGO), width=360)
st.title("⛈️ METAR Extreme Weather Analyser")
st.caption("Downloads a month of METAR/SPECI reports from Ogimet and finds extreme weather events using "
           "**local tiered criteria** and **ICAO Annex 3** criteria. All times are UTC.")

sb = st.sidebar
sb.header("1. Choose station and month")
choice = sb.selectbox("Station", list(STATIONS))
icao = STATIONS[choice]
if icao is None:
    icao = sb.text_input("ICAO code (4 letters)", value="FWKI").strip().upper()
years = list(range(today.year, 1999, -1))
year = sb.selectbox("Year", years, index=years.index(prev.year))
month = sb.selectbox("Month", MONTHS, index=prev.month - 1, format_func=lambda m: pd.Timestamp(2000, m, 1).strftime("%B"))

sb.header("2. Thresholds (optional)")
L = dict(me.LOCAL)
with sb.expander("Local conditions"):
    L["min_event_tier"] = st.selectbox("Count as events from tier", ["Moderate", "Severe", "Very Extreme"],
                                       index=["Moderate", "Severe", "Very Extreme"].index(L["min_event_tier"]))
    L["wind_vext_kt"] = st.number_input("Wind/gust Very Extreme (kt)", value=L["wind_vext_kt"], min_value=1)
    L["gust_spread_kt"] = st.number_input("Gust spread Severe (kt)", value=L["gust_spread_kt"], min_value=1)
    L["vis_moderate_m"] = st.number_input("Visibility Moderate below (m)", value=L["vis_moderate_m"], min_value=1)
    L["vis_severe_m"] = st.number_input("Visibility Severe below (m)", value=L["vis_severe_m"], min_value=1)
    L["vis_vext_m"] = st.number_input("Visibility Very Extreme at/below (m)", value=L["vis_vext_m"], min_value=1)
    L["ceil_moderate_ft"] = st.number_input("Ceiling Moderate at/below (ft)", value=L["ceil_moderate_ft"], min_value=1)
    L["ceil_severe_ft"] = st.number_input("Ceiling Severe below (ft)", value=L["ceil_severe_ft"], min_value=1)
    L["ceil_vext_ft"] = st.number_input("Ceiling Very Extreme at/below (ft)", value=L["ceil_vext_ft"], min_value=1)
    L["heat_moderate_c"] = st.number_input("Heat Moderate (°C)", value=L["heat_moderate_c"])
    L["heat_vext_c"] = st.number_input("Heat Very Extreme (°C)", value=L["heat_vext_c"])
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
                st.session_state["job"] = run_job(icao, year, month, float(gap), L)
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
df, (sec_a, sec_b) = res["df"], res["sections"]
stem = f"{res['icao']}_{res['year']}{res['month']:02d}"
days = pd.Period(f"{res['year']}-{res['month']:02d}").days_in_month

st.subheader(f"{res['icao']} - {pd.Timestamp(res['year'], res['month'], 1).strftime('%B %Y')}")
d1, d2 = st.columns(2)
d1.download_button("⬇️ Download Word report", job["report"], file_name=job["report_name"],
                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
d2.download_button("⬇️ Download everything (ZIP: report, graphs, CSVs)", job["zip"],
                   file_name=f"{stem}_results.zip", mime="application/zip")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Reports analysed", len(df))
m2.metric("Coverage of hourly reports", f"{len(df) / (days * 24) * 100:.0f}%")
m3.metric("Local episodes", len(sec_a["ep"]))
m4.metric("Annex 3 episodes", len(sec_b["ep"]))

tabs = st.tabs(["Overview", "A · Local conditions", "B · ICAO Annex 3", "Graphs"])

with tabs[0]:
    st.markdown("**Key observations**")
    st.dataframe(key_observations(df))
    st.markdown("**Reports in each local tier**")
    tiers = pd.DataFrame({me.CATS[c]: df[f"{c.lower()}_tier"].value_counts() for c in me.CATS}) \
        .reindex(me.TIERS).fillna(0).astype(int).T
    st.dataframe(tiers)

for tab, sec in ((tabs[1], sec_a), (tabs[2], sec_b)):
    with tab:
        st.write(sec["description"])
        ep = sec["ep"]
        if ep.empty:
            st.success("No events detected with these criteria.")
            continue
        ep = ep.assign(level=ep["level"].map(sec["labels"]))
        st.markdown("**Summary by event type**")
        st.dataframe(ep.groupby(["level", "event"]).agg(episodes=("start", "size"), reports=("reports", "sum"),
                                                        first=("start", "min"), last=("end", "max")).reset_index())
        st.markdown("**All episodes**")
        st.dataframe(ep)
        with st.expander("Raw METAR for every flagged report"):
            ev = sec["ev"][["time", "rtype", "event", "level", "raw"]]
            st.dataframe(ev.assign(level=ev["level"].map(sec["labels"])))
        with st.expander("Definitions used"):
            for tag, text in sec["info"].items():
                st.write(f"**{tag}** - {text}")

with tabs[3]:
    for sec in (sec_a, sec_b):
        st.markdown(f"### {sec['title']}")
        for suffix, cap in (("", "Values with event periods shaded"),
                            ("_timeline", "Timeline of every flagged report"),
                            ("_heatmap", "Worst event by day and hour")):
            f = outdir / f"{stem}_{sec['key']}{suffix}.png"
            if f.exists():
                st.image(str(f), caption=cap)

st.caption("Local thresholds are settings that should be agreed with the national meteorological service. "
           "Results depend on report availability; NIL reports are excluded.")
