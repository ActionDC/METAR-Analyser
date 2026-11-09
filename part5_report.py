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


