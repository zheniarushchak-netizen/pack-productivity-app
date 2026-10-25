"""Production interface. Uses existing packing modules without changing formulas."""
import datetime as dt
import hashlib
import html
import time


from io import BytesIO
from packing.demo import demo_frames, as_excel

import pandas as pd
import streamlit as st

from packing.storage import check_storage
from packing.calculations import calculate
from packing.inputs import read_inputs
from packing.excel import export_excel

st.set_page_config(page_title="Packing | Shift report", page_icon="◈", layout="wide")

# Interface styles. Small summary tables use HTML for precise, compact spacing.
st.markdown("""
<style>
.stApp {background:#f4f5f7;color:#202633}
.block-container {max-width:1600px;padding-top:2rem;padding-bottom:3rem}
[data-testid="stSidebar"] {background:#fff;border-right:1px solid #e3e6ec}
h1,h2,h3 {color:#202633;letter-spacing:-.035em}
.brand {display:flex;align-items:center;gap:14px;margin:0 0 26px}
.brand-icon {background:#4f46e5;color:white;padding:12px 16px;border-radius:12px;font-size:26px}
.brand h1 {font-size:32px;margin:0;padding:0;line-height:1.2}
.brand p {font-size:14px;color:#727b8d;margin:5px 0 0}
.kpi {background:white;border:1px solid #e3e6ec;border-radius:14px;padding:18px 20px;margin-bottom:20px}
.kpi-label {font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#687387}
.kpi-value {font-size:clamp(24px,2.7vw,38px);font-weight:700;letter-spacing:-.04em;margin-top:6px;color:#202633}
.kpi-note {font-size:12px;color:#727b8d;margin-top:3px}
.stTabs [data-baseweb="tab-list"] {gap:12px;border-bottom:1px solid #dde2eb}
.stTabs [data-baseweb="tab"] {height:58px;padding:0 18px}
.stTabs [data-baseweb="tab"] p {font-size:18px;font-weight:600}
.stTabs [aria-selected="true"] {color:#4f46e5}
.table-wrap {overflow-x:auto;background:white;border:1px solid #e3e6ec;border-radius:12px;margin:8px 0 22px;max-width:100%}
.compact {border-collapse:collapse;width:max-content;min-width:0;font-variant-numeric:tabular-nums;color:#202633}
.compact th {
    background: #19343D;
    color: #FFFFFF;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .03em;
    white-space: nowrap;
    padding: 10px 13px;
    text-align: right;
}
.compact th:first-child {text-align:left}
.compact td {padding:6px 13px;font-size:18px;font-weight:600;text-align:right;border-top:1px solid #edf0f5;white-space:nowrap;line-height:1.35}
.compact td:first-child {text-align:left;font-size:13px;font-weight:500;color:#626c80}
.compact tr.total td {background:#eef0ff;font-weight:700;color:#3730a3}
.compact tr:hover td {background:#f6f7fc}
.compact td.good {color:#13795b}.compact td.mid {color:#a16207}.compact td.low {color:#b45353}
.empty {background:white;border:1px dashed #c9cfdf;border-radius:18px;padding:42px;margin-top:12px}
.empty h2 {font-size:25px;margin:0 0 10px}.empty p {color:#687387;max-width:650px;line-height:1.7}
@media(max-width:700px){.brand h1{font-size:26px}.compact td{font-size:16px;padding:6px 10px}.stTabs [data-baseweb="tab"]{padding:0 10px}}
 .compact tbody tr.libiao td {
    background: #C5DFE5;
    color: #123B46;
    font-size: 18px;
    font-weight: 800;
    padding: 9px 13px;
    border-bottom: 2px solid #087E8B;
}

.compact tbody tr.libiao td:first-child {
    color: #123B46;
    font-size: 15px;
    font-weight: 800;
    text-transform: uppercase;
} </style>
""", unsafe_allow_html=True)


def number(value, decimals=0, suffix=""):
    if pd.isna(value):
        return "—"
    if isinstance(value, str):
        return value
    return f"{value:,.{decimals}f}{suffix}"


def compact_table(frame, decimals=0, suffix="", rate_colors=False):
    """Render a copy with escaped content. Missing values never become zero."""
    view = frame.copy()
    if isinstance(view.columns, pd.MultiIndex):
        view.columns = [" / ".join(str(x) for x in c if str(x)) for c in view.columns]
    headers = [view.index.name or ""] + list(view.columns)
    out = ['<div class="table-wrap"><table class="compact"><thead><tr>']
    out += [f"<th>{html.escape(str(c))}</th>" for c in headers]
    out.append("</tr></thead><tbody>")
    for idx, row in view.iterrows():
        total = str(idx).lower() == "total"
        if str(idx).strip().lower() == "libiao":
            out.append('<tr class="libiao">')
        elif total:
            out.append('<tr class="total">')
        else:
            out.append('<tr>')
        out.append(f"<td>{html.escape(str(idx))}</td>")
        for value in row:
            color = ""
            if rate_colors and not pd.isna(value) and not isinstance(value, str):
                color = "good" if value >= 95 else "mid" if value >= 80 else "low"
            out.append(f'<td class="{color}">{html.escape(number(value, decimals, suffix))}</td>')
        out.append("</tr>")
    out.append("</tbody></table></div>")
    st.markdown("".join(out), unsafe_allow_html=True)


def data_grid(frame, key):
    """Large reports retain search, sorting and horizontal scrolling."""
    view = frame.copy()
    if isinstance(view.columns, pd.MultiIndex):
        view.columns = [" · ".join(str(x) for x in c if str(x)) for c in view.columns]
    else:
        view.columns = view.columns.map(str)
    view = view.rename(columns={"Work level": "Experience", "Work type": "Work Type",
                                "Last station": "Last Station"})
    config = {}

    for name in view.columns:
        if pd.api.types.is_numeric_dtype(view[name]):
            config[name] = st.column_config.NumberColumn(
                width="small",
                format="%.2f" if any(
                    x in name
                    for x in ["Rate", "UPH", "hours", "gap"]
                ) else "%.0f"
            )

    # Цей блок уже поза циклом for.
    row_styles = pd.DataFrame(
        [
            [
                "background-color: #FFFFFF; color: #19343D;"
                if row_number % 2 == 0
                else "background-color: #E3EFF2; color: #19343D;"
            ] * len(view.columns)
            for row_number in range(len(view))
        ],
        index=view.index,
        columns=view.columns
    )

    styled_view = view.style.apply(
        lambda _: row_styles,
        axis=None
    )

    st.dataframe(
        styled_view,
        width="stretch",
        height=min(620, max(180, 32 * (len(view) + 1))),
        row_height=30,
        column_config=config,
        key=key
    )



def plot(data, x, y, color=None, order=None, height=240, horizontal=False):
    category = {"field": x, "type": "ordinal", "sort": order or None,
                "axis": {"labelAngle": 0 if horizontal else -35, "title": None}}
    measure = {"field": y, "type": "quantitative", "axis": {"title": y, "gridColor":"#e9ecf3"}}
    encoding = {"x": measure if horizontal else category, "y": category if horizontal else measure,
                "tooltip": [{"field": x}, {"field": y, "type":"quantitative", "format":",.0f"}]}
    if color:
        encoding["color"] = {"field":color, "type":"nominal", "scale":{
            "domain":["Single", "Multi"], "range":["#4f46e5", "#a5b4fc"]},
            "legend":{"orient":"top", "title":None}}
        encoding["tooltip"].append({"field":color})
    st.vega_lite_chart(data, {"height":height, "background":"transparent", "mark":{
        "type":"bar", "color":"#6366f1", "cornerRadiusEnd":3}, "encoding":encoding,
        "config":{"view":{"stroke":None}, "axis":{"labelColor":"#687387", "titleColor":"#687387"}}},
        use_container_width=True)


st.markdown('<div class="brand"><div class="brand-icon">◈</div><div><h1>Packing performance</h1>'
            '<p>Shift results · Team productivity · Line output</p></div></div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Report setup")
    st.caption("Upload the two files for one shift.")
    demo_mode = st.toggle("Use demo data", value=True)
    
    if st.button("Перевірити сховище"):
        try:
            check_storage()
            st.success("Підключення працює")
        except Exception:
            st.error(
                "Не вдалося підключитися. "
                "Перевір URL, ключ і назву bucket у Secrets."
            )   
    if demo_mode:
        night = st.toggle("Night shift", value=False)
        extended = st.toggle("Extended shift", value=False)

        demo_pack, demo_employees = demo_frames(
            night=night,
            extended=extended
        )

        pack_file = BytesIO(as_excel(demo_pack))
        pack_file.name = "Demo_operations.xlsx"

        employee_file = BytesIO(as_excel(demo_employees))
        employee_file.name = "Demo_employees.xlsx"

        st.info("Demo data · Fictional operators")

    else:
        pack_file = st.file_uploader(
            "Packing operations",
            type=["xlsx"]
        )

        employee_file = st.file_uploader(
            "Employee directory",
            type=["xlsx"]
        )
    run = st.button("Generate report", type="primary", use_container_width=True,
                    disabled=pack_file is None or employee_file is None)
    st.divider()
    st.caption("Experience is calculated using today's server date, following the original algorithm.")
    with st.expander("Required columns"):
        st.markdown("**Operations:** Package No.; Item quantity classification name; Operator; "
                    "Packing station; Operation time; Quantity of scanned.\n\n"
                    "**Employees:** Employee ID; Create time.")
        st.caption("The first worksheet is read. Operator IDs must appear in parentheses.")

if pack_file is None or employee_file is None:
    st.markdown('<div class="empty"><h2>Your next shift report starts here.</h2>'
                '<p>Upload packing operations and the matching employee directory using the sidebar. '
                'Then select <strong>Generate report</strong> to see actual results.</p>'
                '<p>01 &nbsp; Upload files &nbsp;&nbsp; 02 &nbsp; Calculate &nbsp;&nbsp; 03 &nbsp; Review & export</p></div>',
                unsafe_allow_html=True)
    st.stop()

pack_bytes, employee_bytes = pack_file.getvalue(), employee_file.getvalue()
fingerprint = hashlib.sha256(pack_bytes + employee_bytes + str(dt.date.today()).encode()).hexdigest()

if run:
    st.session_state.pop("packing_result", None)
    st.session_state.pop("packing_excel", None)
    started = time.perf_counter()
    stage = st.empty()
    try:
        with st.spinner("Preparing your report…"):
            stage.info("Step 1 of 2 · Reading Excel files")
            pack, employees = read_inputs(pack_bytes, employee_bytes)
            stage.info("Step 2 of 2 · Calculating shift results")
            result = calculate(pack, employees)
            st.session_state.packing_result = result
            st.session_state.packing_fingerprint = fingerprint
            st.session_state.packing_seconds = time.perf_counter() - started
            st.session_state.packing_names = (pack_file.name, employee_file.name)
        stage.empty()
    except Exception as exc:
        stage.empty()
        st.error("Report generation failed. Check the uploaded files and required columns.")
        with st.expander("Error details"):
            st.code(f"{type(exc).__name__}: {exc}")
        st.stop()

if st.session_state.get("packing_fingerprint") != fingerprint or "packing_result" not in st.session_state:
    st.info("Files are ready. Select Generate report to calculate this shift.")
    st.stop()

result = st.session_state.packing_result
tables = result["tables"]
statistics = tables["statistics_report"].set_index("Metric")["Value"]

heading, export = st.columns([3, 1])

hours = result["hours"]

start_hour = (hours[0] - 1) % 24
end_hour = hours[-1] % 24

report_period = f"{start_hour:02d}:00–{end_hour:02d}:00"

with heading:
    shift_name = "Night shift" if result["night"] else "Day shift"
    extension_label = " · Extended" if result["extended"] else ""

    st.subheader(
        f"{shift_name}{extension_label} · {report_period}"
    )

    # Наявний st.caption(...) залишаємо нижче
    st.caption(f"Calculated {result['calculated_on']} · {len(result['hours'])} reported hours · "
               f"Processing {st.session_state.packing_seconds:.1f}s")
with export:
    # Generate Excel on demand: users can review tables before the export finishes.
    if "packing_excel" not in st.session_state:
        if st.button("Prepare Excel export", use_container_width=True):
            try:
                with st.spinner("Creating Excel workbook…"):
                    st.session_state.packing_excel = export_excel(tables)
                st.rerun()
            except Exception as exc:
                st.error(f"Excel export failed: {exc}")
    else:
        st.download_button("Download Excel", st.session_state.packing_excel, "Pack_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")

for col, label, key, note in zip(st.columns(4), ["Total units", "Single", "Multi", "Operators"],
                               ["Total Quantity", "Total Single", "Total Multi", "Total Operators"],
                               ["All recorded operations", "Single-product units", "Multi-product units", "Original >100-unit rule"]):
    with col:
        st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div>'
                    f'<div class="kpi-value">{html.escape(number(statistics[key]))}</div>'
                    f'<div class="kpi-note">{note}</div></div>', unsafe_allow_html=True)

summary, people, lines, statistics_tab = st.tabs(["Overview", "Operators", "Lines", "Statistics"])
with summary:
    actual_lines = tables["line_report"].drop(index="Manual single", errors="ignore")
    hours = result["hours"]
    labels = [f"{h-1:02}:00–{h%24:02}:00" for h in hours]
    hourly = pd.DataFrame([{"Hour": label, "Type":kind, "Units":actual_lines[(h,kind)].sum()}
                          for h,label in zip(hours,labels) for kind in ["Single", "Multi"]])
    chart_col, zone_col = st.columns([2, 1])
    with chart_col:
        st.subheader("Hourly output")
        plot(hourly, "Hour", "Units", "Type", labels)
    with zone_col:
        st.subheader("Rate distribution")
        zones = tables["rate_zone_report"].set_index("Rate zone").rename(columns={"shape":"Share (%)"})
        compact_table(zones, decimals=2)
    st.subheader("Team performance")
    st.caption("HC, UPH and Rate are shown separately for easier comparison. Values follow the original Dashboard.")
    dash = tables["dashboard_report"]
    for col, metric, title, decimals, suffix in zip(st.columns(3), ["HC", "AVG UPH", "AVG Rate"],
                                                  ["Headcount", "Average UPH", "Average Rate"],
                                                  [0, 2, 2], ["", "", "%"]):
        with col:
            st.markdown(f"**{title}**")
            view = dash.xs(metric, level="Metric").copy()
            view.index.name = "Experience"
            compact_table(view, decimals, suffix, rate_colors=metric=="AVG Rate")
    st.caption("Unknown employees contribute to Veteran / Multi headcount only, following the original rules. "
               "Missing values are shown as —, never as zero.")
    st.subheader("Hourly headcount")
    compact_table(tables["hc_final"])

with people:
        @st.fragment
        def render_people():
            st.subheader("30 найнижчих Rate")

            show_hourly_rate = st.toggle(
            "Показати погодинні дані",
            value=False,
            key="worst_hourly_rate"
            )

            worst_display = tables["worst_rate"]

            columns = [
                ("Operator", ""),
                ("Work level", ""),
                ("Work type", ""),
                ("Last station", ""),
                ("Total", "Rate"),
                ("Total", "Worked hours"),
                ("Total", "UPH"),
                ("Total", "Quantity")
            ]

            if show_hourly_rate:
                for hour in result["hours"]:
                    columns.extend([
                        (hour, "Single"),
                        (hour, "Multi"),
                        (hour, "Total"),
                        (hour, "Rate")
                    ])

            data_grid(
                worst_display.loc[:, columns],
                "worst_rate_grid"
            )
            st.subheader("Operator performance")
            employee = tables["employee_report"]
            search = st.text_input("Find an operator", placeholder="Name or employee ID")
            display = employee.copy()
            if search:
                mask = (display.index.astype(str).str.contains(search, case=False, regex=False)
                        | display[("Operator", "")].astype(str).str.contains(search, case=False, regex=False))
                display = display.loc[mask]
            totals_only = st.toggle("Shift totals only", value=True)
            if totals_only:
                display = display.loc[:, [c for c in display.columns if c[0] in
                                        ["Operator", "Work level", "Work type", "Last station","First scan","Last scan", "Total"]]]
            st.caption(f"Showing {len(display)} of {len(employee)} operators. Search changes this view only; Excel contains the full report.")
            data_grid(display, "operators_grid")
        render_people()
with lines:
    st.subheader("Output by line")
    line_data = actual_lines[("Total", "Quantity")].rename("Units").rename_axis("Line").reset_index()
    line_data = line_data.sort_values("Units", ascending=False)
    plot(line_data, "Line", "Units", order=line_data["Line"].tolist(), horizontal=True,
         height=max(180, min(600, len(line_data)*28)))
    st.caption("The chart excludes the Manual single summary row to avoid counting it twice.")
    st.markdown("**Full line report**")
    data_grid(tables["line_report"], "lines_grid")
    st.subheader("Current headcount by line")
    st.caption("Last available hour in the report.")
    # Розташовуємо номери Libiao під відповідними лініями.
    hc_table = tables["line_hc_report"].copy()

    libiao_row = tables["Libiao"].reindex(
        columns=hc_table.columns
    )

    combined_table = pd.concat([
        libiao_row,
        hc_table
    ])

    compact_table(combined_table)

with statistics_tab:
    st.subheader("Shift statistics")
    a, b = st.columns(2)
    stats = tables["statistics_report"].set_index("Metric")
    with a:
        compact_table(stats.iloc[:7], decimals=2)
    with b:
        compact_table(stats.iloc[7:], decimals=2)
    with st.expander("Report source"):
        st.write("Operations:", st.session_state.packing_names[0])
        st.write("Employees:", st.session_state.packing_names[1])
        st.caption("The calculation functions and original Excel layout are unchanged. "
                   "Formatting and charts are presentation-only. Uploaded data is held in this session.")
