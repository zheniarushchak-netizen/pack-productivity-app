"""Production interface. Uses existing packing modules without changing formulas."""
import datetime as dt
import hashlib
import html
import time


import json
import altair as alt
from io import BytesIO
from packing.demo import demo_frames, as_excel

import pandas as pd
import streamlit as st

from packing.storage import check_storage
from packing.calculations import calculate
from packing.inputs import read_inputs
from packing.excel import export_excel
from packing.inputs import read_inputs, read_employees
from packing.uploads import merge_operations

from packing.storage import (
    check_storage,
    save_employees,
    load_employees,
    save_operations,
    load_operations
)

def show_speed_comparison(table, group_column, title, widget_key):
    st.subheader(title)
    st.caption(
        "Rolling 10-minute output × 6. "
        "Each coloured line represents a separate selection."
    )

    if table is None or table.empty:
        st.info(
            "Generate a report with at least 10 complete minutes "
            "of data to display this chart."
        )
        return

    options = sorted(table[group_column].dropna().unique().tolist())

    if group_column == "Process":
        preferred_order = ["Multi", "Manual Single", "PM", "Unknown"]
        options = [
            name for name in preferred_order if name in options
        ]
        default = options
        label = "Processes"
    else:
        default = options[:3]
        label = "Packing lines"

    selected = st.multiselect(
        label,
        options=options,
        default=default,
        key=widget_key,
    )

    if not selected:
        st.info("Select at least one option to display the chart.")
        return

    chart_data = (
        table.loc[table[group_column].isin(selected)]
        .sort_values(["Time", group_column])
    )

    # Inline data also supports reports with many time points.
    chart_values = json.loads(
        chart_data.to_json(orient="records", date_format="iso")
    )

    chart = (
        alt.Chart(alt.Data(values=chart_values))
        .mark_line(strokeWidth=2.8, clip=True)
        .encode(
            x=alt.X(
                "Time:T",
                title=None,
                axis=alt.Axis(
                    format="%H:%M",
                    tickCount=12,
                    labelAngle=0,
                    grid=True,
                ),
            ),
            y=alt.Y(
                "Units per hour:Q",
                title="Units / hour",
                scale=alt.Scale(zero=True),
                axis=alt.Axis(format="~s", tickCount=6),
            ),
            color=alt.Color(
                f"{group_column}:N",
                title=None,
                scale=alt.Scale(
                    domain=options,
                    range=[
                        "#176b50",
                        "#10B981",
                        "#F59E0B",
                        "#EF4444",
                        "#0EA5E9",
                        "#EC4899",
                        "#14B8A6",
                        "#8B5CF6",
                        "#84CC16",
                        "#F97316",
                        "#64748B",
                        "#A855F7",
                        "#06B6D4",
                        "#D97706",
                        "#BE185D",
                        "#047857",
                    ],
                ),
                legend=alt.Legend(
                    orient="top",
                    columns=4,
                    labelFontSize=12,
                    symbolStrokeWidth=3,
                ),
            ),
            tooltip=[
                alt.Tooltip(f"{group_column}:N", title=group_column),
                alt.Tooltip(
                    "Time:T",
                    title="Window ends",
                    format="%H:%M",
                ),
                alt.Tooltip(
                    "Units per hour:Q",
                    title="Units / hour",
                    format=",.0f",
                ),
            ],
        )
        .properties(
            height=560,
            background="transparent",
            padding={"left": 5, "right": 20, "top": 15, "bottom": 10},
        )
        .configure_view(stroke=None)
        .configure_axis(
            labelFontSize=12,
            titleFontSize=12,
            labelColor="#64748B",
            titleColor="#64748B",
            gridColor="#E2E8F0",
            gridOpacity=0.7,
            domain=False,
            tickSize=0,
            labelPadding=10,
            titlePadding=15,
        )
    )

    st.altair_chart(chart, use_container_width=True, theme=None)

st.set_page_config(page_title="Packing | Shift report", page_icon="◈", layout="wide", initial_sidebar_state="collapsed")

# Presentation only: emerald / warm-white theme.
st.markdown("""
<style>
.stApp {background:#faf8f3;color:#182e29}
.block-container {max-width:1760px;padding-top:1.5rem;padding-bottom:4rem}
[data-testid="stSidebar"] {background:#fff;border-right:1px solid #e0e8e2}
h1,h2,h3 {color:#183c31;letter-spacing:-.035em}
.brand {display:flex;align-items:center;gap:16px;margin-bottom:26px}
.brand-icon {background:#176b50;color:#fff;padding:12px 18px;border-radius:16px;font-size:28px}
.brand h1 {font-size:34px;line-height:1.2;margin:0;padding:0}
.brand p,.kpi-note {color:#73847b;font-size:13px;margin-top:7px}
.kpi {background:#fff;border:1px solid #e2e8df;border-radius:20px;padding:24px;margin-bottom:16px;min-height:146px}
.kpi-label {font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:#6b7e73;font-weight:600}
.kpi-value {font-size:clamp(25px,3vw,42px);font-weight:700;letter-spacing:-.05em;color:#183c31;margin-top:10px}
.stTabs [data-baseweb="tab-list"] {gap:26px;border-bottom:1px solid #dfe7df;margin-bottom:22px}
.stTabs [data-baseweb="tab"] {height:54px;padding:0 8px}
.stTabs [aria-selected="true"] {color:#176b50!important}
.stTabs [data-baseweb="tab-highlight"] {background:#176b50}
button[kind="primary"] {background:#176b50!important;border-color:#176b50!important;border-radius:12px}
[data-testid="stVegaLiteChart"] {background:#fff;border:1px solid #e2e8df;border-radius:18px;padding:16px}
.table-wrap {overflow-x:auto;background:#fff;border:1px solid #e2e8df;border-radius:16px;margin:12px 0 24px}
.compact {border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}
.compact th {background:#edf3ed;color:#456353;text-transform:uppercase;font-size:11px;letter-spacing:.04em;padding:12px;text-align:right;white-space:nowrap}
.compact td {padding:10px 12px;text-align:right;border-top:1px solid #edf1eb;white-space:nowrap;font-size:15px}
.compact td:first-child,.compact th:first-child {text-align:left}
.compact tr.total td {background:#eaf4ed;color:#176b50;font-weight:700}
.compact tr:hover td {background:#f5f8f3}
.compact td.good {color:#187c54}.compact td.mid {color:#ad7828}.compact td.low {color:#bd6759}
.compact tr.libiao td {background:#deeee2;color:#176b50;font-weight:700}
.empty {padding:48px;background:white;border:1px dashed #c6d9cb;border-radius:22px}
.performance-card {background:white;border:1px solid #e2e8df;border-radius:22px;padding:24px;margin-bottom:24px;min-height:225px}
.performance-card.hero {background:#183e32;color:white}
.performance-card.hero .kpi-label,.performance-card.hero .ring-note {color:#c3d7cb}
.ring-row {display:flex;align-items:center;gap:14px;margin-top:18px;flex-wrap:wrap;align-items:center}
.ring {width:122px;height:122px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.ring-center {background:white;width:100px;height:100px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:25px;font-weight:700;color:#183c31}
.hero .ring-center {background:#183e32;color:white}
.ring-note {font-size:12px;color:#73847b;line-height:1.8}
@media(max-width:700px){.block-container{padding-top:1rem}.brand h1{font-size:26px}.stTabs [data-baseweb="tab-list"]{gap:8px}.kpi{padding:18px}}

.summary-hero{display:flex;align-items:center;gap:30px;border:1px solid #dfdfd4;border-radius:20px;background:#fffdf9;padding:24px;margin-bottom:22px;flex-wrap:wrap}
.large-ring{width:170px;height:170px;border-radius:50%;display:flex;align-items:center;justify-content:center;margin-top:14px}
.large-ring>div{background:#fffdf9;width:140px;height:140px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:36px;font-weight:750;color:#15392a}
.hero-output{flex:1;min-width:190px}.hero-output>div:not(.kpi-value):not(.kpi-label){line-height:2}.hero-output b{float:right}.hero-output hr{border:0;border-top:1px solid #e5e7dd}
.experience-panel{padding:24px;background:#fffdf9;border:1px solid #dfdfd4;border-radius:18px}.composition{display:flex;overflow:hidden;border-radius:9px;height:44px}.composition>div{display:flex;align-items:center;justify-content:center;color:white;font-weight:700;min-width:0;font-size:13px}.composition-legend{display:flex;flex-wrap:wrap;justify-content:space-between;gap:14px;margin-top:22px;font-size:13px}.composition-legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px}
.stTabs [data-baseweb="tab"]{border-radius:10px;padding:0 22px;height:46px}.stTabs [aria-selected="true"]{background:#176b50;color:white!important}
.brand h1{font-family:Georgia,serif}.brand{margin-bottom:12px}.brand-icon{padding:8px 14px}

.report-card{background:#fffdf9;border:1px solid #e5ded2;border-radius:18px;padding:16px 12px;margin:12px 0 18px;box-shadow:0 3px 12px #183c3105}
.report-card header{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:0 6px 16px;flex-wrap:wrap}.report-card h3{font-family:Georgia,serif;font-size:21px!important;margin:0;padding:0}.report-card header p{font-size:11px;color:#778278;margin:5px 0 0}
.card-total{background:#176b50;color:white;border-radius:9px;padding:6px 11px;font-size:20px;font-weight:700;font-variant-numeric:tabular-nums}
.report-scroll{overflow-x:auto}.report-table{border-collapse:separate;border-spacing:0;width:100%;font-variant-numeric:tabular-nums;font-size:13px;color:#17382e}
.report-table th,.report-table td{border:0!important;padding:10px 7px;text-align:right;white-space:nowrap}.report-table thead th{background:#f5f1e9;font-size:10px;text-transform:uppercase;letter-spacing:.02em}.report-table th:first-child{text-align:left}.report-table tbody th{font-weight:500}.report-table tbody td,.report-table tbody th{border-bottom:1px solid #eeeee6!important}
.report-table .total-col{background:#f0f5e9}.report-table .total-row td,.report-table .total-row th{background:#eaf3e4;font-weight:750}.report-table tbody tr:hover{outline:1px solid #d5e3d6;outline-offset:-1px}
.rate-pill{display:inline-block;padding:3px 6px;border-radius:7px;font-weight:650;font-size:12px}.rate-pill.good{background:#dcf4df;color:#17633e}.rate-pill.amber{background:#fff0c8;color:#8b5c10}.rate-pill.orange{background:#ffe5ce;color:#a54a19}.rate-pill.red{background:#ffe0dc;color:#ab3730}.rate-pill.neutral{background:#f3f0e9;color:#6b756e}

/* Readable, higher-contrast report tables. */
.report-card{background:#fff;border:1px solid #cbd8cb;border-top:3px solid #2b7956;box-shadow:0 5px 16px rgba(24,60,49,.07);padding:18px 12px}
.report-card header{min-height:74px;padding-bottom:16px}.report-card h3{font-size:24px!important;color:#123e2c}.report-card header p{font-size:12px;color:#526659}
.report-card header>.rate-pill,.card-total{font-size:23px;line-height:1.3;padding:8px 12px;border-radius:10px;font-weight:750}
.report-table{font-size:17px;color:#153a2b}
.report-table th,.report-table td{padding:12px 7px;height:52px;box-sizing:border-box}
.report-table td{font-weight:600}
.report-table thead th{font-size:11px;color:#294f3b;background:#e6eee3;letter-spacing:.025em;height:43px}
.report-table tbody th{font-size:15px;font-weight:600}
.report-table .total-col{background:#e5f0da;font-weight:700}
.report-table .total-row td,.report-table .total-row th{background:#d2e8c6;color:#124b30;font-weight:800;border-top:2px solid #9fc195!important}
.report-table .rate-pill{font-size:15px;line-height:1.4;padding:4px 6px;font-weight:750}
.rate-pill.good{background:#c6edcd;color:#10582e}.rate-pill.amber{background:#ffe4a0;color:#774500}.rate-pill.orange{background:#ffd2b5;color:#963b0b}.rate-pill.red{background:#ffc7c1;color:#98251e}
@media(max-width:1100px){.report-table{min-width:410px}}

.performance-card{padding:18px 14px;min-height:270px;margin-top:12px;margin-bottom:18px}
.performance-card .kpi-label{font-size:14px;font-weight:750;letter-spacing:.035em;color:#234a36}
.performance-card .ring-row{display:flex;flex-direction:column;align-items:center;gap:12px;margin-top:14px;flex-wrap:nowrap}
.performance-card .ring{width:104px;height:104px}.performance-card .ring-center{width:84px;height:84px;font-size:26px}
.performance-card .ring-note{font-size:15px;line-height:1.7;color:#415b4c;text-align:center}.performance-card .ring-note b{font-size:19px;color:#183e2b}
.distribution-card{padding:12px;margin-top:12px;min-height:270px}.distribution-card .report-table td{font-size:17px}.distribution-card .report-table th,.distribution-card .report-table td{height:49px;padding:9px 8px}
@media(max-width:1100px){.distribution-card .report-table{min-width:300px}}
</style>
""", unsafe_allow_html=True)


def number(value, decimals=0, suffix=""):
    if pd.isna(value):
        return "—"
    if isinstance(value, str):
        return value
    return f"{value:,.{decimals}f}{suffix}"


def performance_card(label, rate, uph, hc, hero=False):
    value = pd.to_numeric(rate, errors="coerce")
    progress = 0 if pd.isna(value) else min(100, max(0, float(value)))
    colour = "#94a39a" if pd.isna(value) else ("#278561" if value >= 95 else "#c69b38" if value >= 80 else "#d98b48" if value >= 60 else "#d16b61")
    track = "#315747" if hero else "#edf2e9"
    text = "—" if pd.isna(value) else f"{value:.1f}%"
    kind = "performance-card hero" if hero else "performance-card"
    st.markdown(
        f'<div class="{kind}"><div class="kpi-label">{html.escape(label)}</div>'
        f'<div class="ring-row"><div class="ring" style="background:conic-gradient({colour} {progress}%, {track} 0)">'
        f'<div class="ring-center">{text}</div></div><div class="ring-note">'
        f'Average rate<br><b>{html.escape(number(uph, 1))}</b> UPH<br>'
        f'<b>{html.escape(number(hc, 2))}</b> equivalent HC</div></div></div>',
        unsafe_allow_html=True,
    )


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


def rate_badge(value):
    numeric = pd.to_numeric(value, errors="coerce")
    if pd.isna(numeric):
        return '<span class="rate-pill neutral">—</span>'
    zone = "good" if numeric >= 95 else "amber" if numeric >= 80 else "orange" if numeric >= 60 else "red"
    return f'<span class="rate-pill {zone}">{numeric:.2f}%</span>'


def report_card(frame, title, subtitle, rate=False, hourly=False):
    """Presentation only; original values and row order are preserved."""
    esc = lambda value: html.escape(str(value))
    total = frame.loc["Total", "Total"] if "Total" in frame.index and "Total" in frame.columns else None
    badge = ""
    if total is not None:
        badge = rate_badge(total) if rate else f'<span class="card-total">{esc(number(total, 2))}</span>'
    out = [f'<section class="report-card"><header><div><h3>{esc(title)}</h3><p>{esc(subtitle)}</p></div>{badge}</header>',
           '<div class="report-scroll"><table class="report-table"><thead><tr>',
           f'<th scope="col">{esc(frame.index.name or "Experience")}</th>']
    for column in frame.columns:
        cls = ' class="total-col"' if str(column) == "Total" else ""
        out.append(f'<th scope="col"{cls}>{esc(column)}</th>')
    out.append('</tr></thead><tbody>')
    for label, row in frame.iterrows():
        cls = ' class="total-row"' if str(label) == "Total" else ""
        out.append(f'<tr{cls}><th scope="row">{esc(label)}</th>')
        nums = pd.to_numeric(row, errors="coerce")
        low, high = nums.min(), nums.max()
        for column, value in row.items():
            cell_cls = ' class="total-col"' if str(column) == "Total" else ""
            style = ""
            numeric = pd.to_numeric(value, errors="coerce")
            if hourly and pd.notna(numeric):
                strength = 0.5 if high == low else (numeric - low) / (high - low)
                alpha = 0.10 + 0.30 * strength
                style = f' style="background:rgba(39,133,81,{alpha:.3f})"'
            content = rate_badge(value) if rate else esc(number(value, 0 if hourly else 2))
            out.append(f'<td{cell_cls}{style}>{content}</td>')
        out.append('</tr>')
    out.append('</tbody></table></div></section>')
    st.markdown(''.join(out), unsafe_allow_html=True)


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
                else "background-color: #F0F5EE; color: #19343D;"
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
            "domain":["Single", "Multi"], "range":["#d99b7b", "#85b89b"]},
            "legend":{"orient":"top", "title":None}}
        encoding["tooltip"].append({"field":color})
    st.vega_lite_chart(data, {"height":height, "background":"transparent", "mark":{
        "type":"bar", "color":"#278561", "cornerRadiusEnd":3}, "encoding":encoding,
        "config":{"view":{"stroke":None}, "axis":{"labelColor":"#687387", "titleColor":"#687387"}}},
        use_container_width=True)


st.markdown('<div class="brand"><div class="brand-icon">◈</div><div><h1>Packing Analytics</h1>'
            '<p>SHIFT INTELLIGENCE / Output · People · Lines</p></div></div>', unsafe_allow_html=True)

with st.sidebar:

    st.header("Report setup")

    st.caption("Upload the two files for one shift.")

    demo_mode = st.toggle("Use demo data", value=True)
    
    if st.button("Перевірити сховище"):
        try:
            check_storage()
            st.success("Connection is ready")
        except Exception as exc:
            st.error(
                f"Не вдалося підключитися. "
                f"Тип помилки: {type(exc).__name__}"
            )
    with st.expander("Оновити довідник працівників"):

        new_employee_file = st.file_uploader(
            "Новий довідник",
            type=["xlsx"],
            key="employee_directory_update"
        )

        if st.button(
            "Зберегти довідник",
            disabled=new_employee_file is None,
            key="save_employee_directory"
        ):
            stage = "Читання завантаженого файлу"

            try:
                employee_content = new_employee_file.getvalue()

                stage = "Перевірка Excel"
                checked_employees = read_employees(employee_content)

                stage = "Запис у сховище"
                save_employees(employee_content)

                st.success(
                    f"Довідник збережено. "
                    f"Рядків: {len(checked_employees)}"
                )

            except Exception as exc:
                st.error(
                    f"Етап: {stage}. "
                    f"Тип помилки: {type(exc).__name__}"
                )
                

                if isinstance(exc, ValueError):
                    st.write(str(exc))
    if demo_mode:
        # Generate demo files
        demo_pack, demo_employees = demo_frames(
            night=False,
            extended=False
        )

        pack_file = BytesIO(as_excel(demo_pack))
        pack_file.name = "Demo_operations.xlsx"

        employee_file = BytesIO(as_excel(demo_employees))
        employee_file.name = "Demo_employees.xlsx"

    else:
        # Upload real operations
        pack_file = st.file_uploader(
            "Packing operations",
            type=["xlsx"]
        )

        # Load the saved employee directory
        employee_file = None

        try:
            saved_employee_bytes = load_employees()

            employee_file = BytesIO(saved_employee_bytes)
            employee_file.name = "Saved_employee_directory.xlsx"

            st.caption("Using the saved employee directory.")

        except Exception as exc:
            st.warning(
                "Employee directory unavailable. "
                "Upload it using 'Update employee directory'. "
                f"Error type: {type(exc).__name__}"
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
    started = time.perf_counter()
    stage = st.empty()

    try:
        with st.spinner("Preparing your report…"):
            stage.info("Reading Excel files")
            uploaded_operations, employees = read_inputs(
                pack_bytes,
                employee_bytes
            )

            if demo_mode:
                operations = uploaded_operations
            else:
                saved_operations = st.session_state.get(
                    "saved_operations"
                )
                loaded_until = st.session_state.get(
                    "loaded_until"
                )
                file_times = pd.to_datetime(
                    uploaded_operations["Operation time"],
                    errors="raise"
                )

                file_start = file_times.min()
                file_end = file_times.max()
                operations, added_operations, updated_loaded_until = (
                    merge_operations(
                        saved_operations=saved_operations,
                        uploaded_operations=uploaded_operations,
                        loaded_until=loaded_until
                    )
                )

            stage.info("Calculating shift results")
            result = calculate(
                        operations,
                        employees,
                        report_work_hours=6.25
                    )

            # Update saved data only after successful calculation
            if not demo_mode:
                st.session_state.saved_operations = operations
                st.session_state.loaded_until = updated_loaded_until

                history = st.session_state.get("upload_history", [])

                st.session_state.upload_history = history + [{
                    "File": pack_file.name,
                    "File start": (
                        file_start.strftime("%Y-%m-%d %H:%M:%S")
                        if pd.notna(file_start) else "—"
                    ),
                    "File end": (
                        file_end.strftime("%Y-%m-%d %H:%M:%S")
                        if pd.notna(file_end) else "—"
                    ),
                    "Added rows": len(added_operations),
                    "Last loaded scan": updated_loaded_until
                }]

            st.session_state.packing_result = result
            st.session_state.packing_fingerprint = fingerprint
            st.session_state.packing_seconds = (
                time.perf_counter() - started
            )
            st.session_state.packing_names = (
                pack_file.name,
                employee_file.name
            )

            # Remove the export belonging to the previous report
            st.session_state.pop("packing_excel", None)

        stage.empty()

    except Exception as exc:
        stage.empty()
        st.error("Report generation failed.")

        with st.expander("Error details"):
            st.exception(exc)

        st.stop()

if st.session_state.get("packing_fingerprint") != fingerprint or "packing_result" not in st.session_state:
    st.info("Files are ready. Select Generate report to calculate this shift.")
    st.stop()

result = st.session_state.packing_result
if not demo_mode:
    loaded_until = st.session_state.get("loaded_until")

    if loaded_until is not None:
        st.caption(
            f"Last loaded scan: {loaded_until:%Y-%m-%d %H:%M:%S}"
        )

    with st.sidebar.expander("Upload history", expanded=False):
        if st.button("Clear shift data", use_container_width=True):
            for key in [
                "saved_operations",
                "loaded_until",
                "upload_history",
                "packing_result",
                "packing_fingerprint",
                "packing_seconds",
                "packing_names",
                "packing_excel",
            ]:
                st.session_state.pop(key, None)

            st.rerun()
        history = st.session_state.get("upload_history", [])

        if history:
            st.dataframe(
                pd.DataFrame(history),
                use_container_width=True,
                hide_index=True
            )
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

summary, people, lines, statistics_tab = st.tabs(["Overview", "Operators", "Lines", "Reports"])
with summary:
    dashboard = tables["dashboard_report"]
    rate_view = dashboard.xs("AVG Rate", level="Metric")
    uph_view = dashboard.xs("AVG UPH", level="Metric")
    hc_view = dashboard.xs("HC", level="Metric")
    def dashboard_total(frame, column):
        return frame.loc["Total", column] if "Total" in frame.index and column in frame.columns else float("nan")
    overview_left, overview_right = st.columns([1.15, 1])
    with overview_left:
        overall = pd.to_numeric(dashboard_total(rate_view, "Total"), errors="coerce")
        fill = 0 if pd.isna(overall) else min(100, max(0, float(overall)))
        st.markdown(
            '<div class="summary-hero"><div><div class="kpi-label">Overall performance</div>'
            f'<div class="large-ring" style="background:conic-gradient(#178451 {fill}%,#e6ebe1 0)">'
            f'<div>{number(overall, 1, "%")}</div></div></div><div class="hero-output">'
            '<div class="kpi-label">Shift output</div>'
            f'<div class="kpi-value">{number(statistics.get("Total Quantity", float("nan")))}</div>'
            '<div class="ring-note">units packed</div><hr>'
            f'<div>Equivalent HC <b>{number(dashboard_total(hc_view, "Total"), 2)}</b></div>'
            f'<div>Average UPH <b>{number(dashboard_total(uph_view, "Total"), 1)}</b></div>'
            f'<div>Shift <b>{html.escape(report_period)}</b></div></div></div>',
            unsafe_allow_html=True,
        )
    with overview_right:
        st.markdown("#### Experience composition")
        experience = hc_view.drop(index="Total", errors="ignore")["Total"]
        experience = pd.to_numeric(experience, errors="coerce").dropna()
        experience = experience[experience > 0]
        total_hc = experience.sum()
        colours = ["#176b50", "#4caa76", "#9ccd9b", "#e39b84"]
        if total_hc > 0:
            segments = []
            legend = []
            for i, (level, value) in enumerate(experience.items()):
                share = value / total_hc * 100
                colour = colours[i % len(colours)]
                label = html.escape(str(level))
                segments.append(f'<div style="flex:{value};background:{colour}" title="{label}: {share:.1f}%">{share:.0f}%</div>')
                legend.append(f'<span><i style="background:{colour}"></i>{label}<br><b>{value:.2f} HC</b></span>')
            st.markdown('<div class="experience-panel"><div class="composition">'+''.join(segments)+'</div><div class="composition-legend">'+''.join(legend)+'</div></div>', unsafe_allow_html=True)
        else:
            st.info("Experience composition is unavailable for this report.")
        st.caption("Share of equivalent headcount, based on the report's experience groups.")
    actual_lines = tables["line_report"].drop(index="Manual single", errors="ignore")
    hours = result["hours"]
    labels = [f"{h-1:02}:00–{h%24:02}:00" for h in hours]
    hourly = pd.DataFrame([{"Hour": label, "Type":kind, "Units":actual_lines[(h,kind)].sum()}
                          for h,label in zip(hours,labels) for kind in ["Single", "Multi"]])
    process_area, distribution_area = st.columns([2.2, 1])
    with process_area:
        st.subheader("Process performance")
        for col, label, process in zip(st.columns(3),
                                       ["Multi", "Manual Single", "Pack Machine"],
                                       ["Multi", "Single", "PM"]):
            with col:
                performance_card(label, dashboard_total(rate_view, process),
                                 dashboard_total(uph_view, process), dashboard_total(hc_view, process))
    with distribution_area:
        st.subheader("Rate distribution")
        zones = tables["rate_zone_report"].set_index("Rate zone").rename(columns={"shape": "Share (%)"})
        zone_rows = []
        for i, (zone, row) in enumerate(zones.iterrows()):
            zone_class = ["good", "amber", "orange", "red"][min(i, 3)]
            cells = []
            for column, value in row.items():
                is_share = "share" in str(column).lower()
                cells.append(f'<td>{html.escape(number(value, 2 if is_share else 0, "%" if is_share else ""))}</td>')
            zone_rows.append(f'<tr><th scope="row"><span class="rate-pill {zone_class}">{html.escape(str(zone))}</span></th>'+''.join(cells)+'</tr>')
        headers = ''.join(f'<th scope="col">{html.escape(str(c))}</th>' for c in zones.columns)
        st.markdown('<section class="report-card distribution-card"><div class="report-scroll"><table class="report-table"><thead><tr><th scope="col">Rate zone</th>'+headers+'</tr></thead><tbody>'+''.join(zone_rows)+'</tbody></table></div></section>', unsafe_allow_html=True)

    st.subheader("Packing speed")
    st.caption(
        "Total speed is always displayed. "
        "Select processes to compare with the total. "
        "Speed = output over the preceding 10 minutes × 6."
    )

    filter_pm, filter_single, filter_multi, _ = st.columns([1, 1.4, 1, 5])

    with filter_pm:
        show_pm = st.checkbox("PM", key="speed_show_pm")

    with filter_single:
        show_single = st.checkbox(
            "Manual Single",
            key="speed_show_single",
        )

    with filter_multi:
        show_multi = st.checkbox("Multi", key="speed_show_multi")

    speed = tables.get("packing_speed")
    process_speed = tables.get("packing_speed_by_process")

    if speed is None:
        st.info("Packing speed is missing from this report. Regenerate the report and check the calculation output.")
    elif speed.empty:
        st.info("No complete 10-minute speed windows are available in this report.")
    else:
        total = speed[["Time", "Units per hour"]].copy()
        total["Series"] = "Total"

        selected_processes = []

        if show_pm:
            selected_processes.append("PM")

        if show_single:
            selected_processes.append("Manual Single")

        if show_multi:
            selected_processes.append("Multi")

        parts = [total]

        if selected_processes:
            if process_speed is None or process_speed.empty:
                st.info(
                    "Generate the report again to load process speeds."
                )
            else:
                selected_data = (
                    process_speed.loc[
                        process_speed["Process"].isin(selected_processes),
                        ["Time", "Process", "Units per hour"],
                    ]
                    .rename(columns={"Process": "Series"})
                )

                parts.append(selected_data)

        chart_data = (
            pd.concat(parts, ignore_index=True)
            .sort_values(["Time", "Series"])
        )

        chart_data["Time"] = pd.to_datetime(chart_data["Time"])
        available_times = sorted(chart_data["Time"].dropna().unique())
        if len(available_times) > 1:
            window = st.select_slider(
                "Visible time range",
                options=available_times,
                value=(available_times[0], available_times[-1]),
                format_func=lambda value: pd.Timestamp(value).strftime("%H:%M"),
                key="speed_visible_window",
            )
            chart_data = chart_data.loc[chart_data["Time"].between(*window)]
        st.caption("Drag the range handles to inspect an interval. Scroll over the chart to zoom; double-click to reset zoom. Times mark the end of each 10-minute window.")
        chart_values = json.loads(
            chart_data.to_json(
                orient="records",
                date_format="iso",
            )
        )

        base = alt.Chart(
            alt.Data(values=chart_values)
        ).encode(
            x=alt.X(
                "Time:T",
                title=None,
                axis=alt.Axis(
                    format="%H:%M",
                    tickCount=12,
                    labelAngle=0,
                    grid=True,
                ),
            ),
            y=alt.Y(
                "Units per hour:Q",
                title="Units / hour",
                scale=alt.Scale(zero=True),
                axis=alt.Axis(format="~s", tickCount=6),
            ),
        )

        # Light shading only under the total curve.
        area = (
            base
            .transform_filter(alt.datum.Series == "Total")
            .mark_area(
                color="#176b50",
                opacity=0.08,
            )
        )

        curves = base.mark_line(
            strokeWidth=3,
            clip=True,
        ).encode(
            color=alt.Color(
                "Series:N",
                title=None,
                scale=alt.Scale(
                    domain=["Total", "PM", "Manual Single", "Multi"],
                    range=["#176b50", "#F59E0B", "#0EA5E9", "#10B981"],
                ),
                legend=alt.Legend(
                    orient="top",
                    labelFontSize=12,
                    symbolStrokeWidth=3,
                ),
            ),
            tooltip=[
                alt.Tooltip("Series:N", title="Process"),
                alt.Tooltip(
                    "Time:T",
                    title="Window ends",
                    format="%H:%M",
                ),
                alt.Tooltip(
                    "Units per hour:Q",
                    title="Units / hour",
                    format=",.0f",
                ),
            ],
        )

        chart = (
            (area + curves)
            .interactive(bind_y=False)
            .properties(
                height=560,
                background="transparent",
                padding={
                    "left": 5,
                    "right": 20,
                    "top": 15,
                    "bottom": 10,
                },
            )
            .configure_view(stroke=None)
            .configure_axis(
                labelFontSize=12,
                titleFontSize=12,
                labelColor="#64748B",
                titleColor="#64748B",
                gridColor="#E2E8F0",
                gridOpacity=0.7,
                domain=False,
                tickSize=0,
                labelPadding=10,
                titlePadding=15,
            )
        )

        st.altair_chart(
            chart,
            use_container_width=True,
            theme=None,
        )
    st.subheader("Hourly output")
    plot(hourly, "Hour", "Units", "Type", labels, height=300)

    st.subheader("Team performance")
    st.caption("Experience and process comparison · Exact report values")
    dash = tables["dashboard_report"]
    for col, metric, title, subtitle in zip(
        st.columns(3), ["HC", "AVG UPH", "AVG Rate"],
        ["Equivalent HC", "Average UPH", "Average Rate"],
        ["Equivalent headcount", "Units per worked hour", "Performance against the assigned norm"],
    ):
        with col:
            view = dash.xs(metric, level="Metric").copy()
            view.index.name = "Experience"
            report_card(view, title, subtitle, rate=metric == "AVG Rate")
    st.caption("Missing values are shown as —. Rate zones: ≥95% · 80–<95% · 60–<80% · <60%.")
    hc_display = tables["hc_final"].copy()
    hc_display.columns = [f"{(int(hour) - 1) % 24:02d}:00" for hour in hc_display.columns]
    hc_display.index.name = "Process"
    report_card(hc_display, "Hourly headcount", f"{len(hc_display.columns)} hourly intervals", hourly=True)
    st.caption("Shading compares hours within each row; it does not indicate performance. Hour labels mark the start of the interval.")

with people:
    @st.fragment
    def render_people():
        employee = tables["employee_report"]
        def source(column):
            return employee[column] if column in employee.columns else pd.Series(pd.NA, index=employee.index)
        view = pd.DataFrame(index=employee.index)
        view["Operator"] = source(("Operator", ""))
        view["Experience"] = source(("Work level", "")).fillna("Unknown")
        view["Process"] = source(("Work type", "")).replace({"Single": "Manual Single"}).fillna("Unknown")
        view["Last station"] = source(("Last station", ""))
        def scan_text(value):
            if pd.isna(value):
                return "—"
            if hasattr(value, "strftime"):
                return value.strftime("%H:%M:%S")
            return str(value)
        view["First scan"] = source(("First scan", "")).map(scan_text)
        view["Last scan"] = source(("Last scan", "")).map(scan_text)
        # Canonical field returned by calculate(); 0 means no gap detected.
        break_column = ("Total", "Break gap min")
        if break_column in employee.columns:
            break_minutes = pd.to_numeric(employee[break_column], errors="coerce")
            view["Break time"] = break_minutes.where(break_minutes.gt(0))
        else:
            break_column = None
            view["Break time"] = float("nan")
        for label, field in [("Units", "Quantity"), ("Hours", "Worked hours"), ("UPH", "UPH"), ("Rate", "Rate")]:
            view[label] = pd.to_numeric(source(("Total", field)), errors="coerce")
        def last_line(station):
            if pd.isna(station) or not str(station).strip():
                return "Unknown"
            station = str(station).strip().upper()
            if station.startswith("PLA"):
                return "PM"
            if station.startswith(("PLE", "PL0")):
                return "LINE 0"
            return station.split("-")[0]
        view["Last line"] = view["Last station"].map(last_line)
        view.index.name = "Employee ID"
        st.subheader("Operator performance")
        st.caption("Search, compare and inspect individual results. Filters affect this view only.")
        def reset_operator_filters():
            for key, value in {"op_line": "All", "op_search": "", "op_process": "All", "op_experience": "All", "op_zone": "All", "op_mode": "All operators"}.items():
                st.session_state[key] = value
        st.markdown("""
        <style>
        .st-key-operator_filters {
            background:#ffffff; border:1px solid #d5e1d9;
            border-radius:14px; padding:12px 16px;
            box-shadow:0 2px 8px rgba(21,61,48,.04);
        }
        .st-key-operator_shift_table {
            border:2px solid #9ebcaf; border-top:4px solid #176b50;
            border-radius:14px; padding:6px; background:#ffffff;
            box-shadow:0 3px 12px rgba(21,61,48,.07);
        }
        .st-key-operator_view_switch button {
            min-height:54px; border-radius:12px; font-weight:750;
        }
        .st-key-operator_view_switch button p {font-size:18px;font-weight:750;}
        .st-key-operator_filters [data-testid="stVerticalBlock"] {gap:0.45rem;}
        </style>
        """, unsafe_allow_html=True)
        if "op_display" not in st.session_state:
            st.session_state["op_display"] = "Shift totals"
        def change_operator_view(value):
            st.session_state["op_display"] = value
        with st.container(key="operator_view_switch"):
            totals_button, hourly_button = st.columns(2)
            for col, label in [(totals_button, "Shift totals"), (hourly_button, "Hourly")]:
                with col:
                    st.button(label, key=f"op_switch_{label}", use_container_width=True,
                              type="primary" if st.session_state["op_display"] == label else "secondary",
                              on_click=change_operator_view, args=(label,))
        display = st.session_state["op_display"]
        with st.container(key="operator_filters"):
            search_col, mode_col, reset_col = st.columns([3, 1.3, 0.8])
            with search_col:
                search = st.text_input("Search operators", placeholder="Search name or login…",
                                       key="op_search", label_visibility="collapsed")
            with mode_col:
                mode = st.selectbox("Operator list", ["All operators", "Lowest 30"],
                                    key="op_mode", label_visibility="collapsed")
            with reset_col:
                st.button("Reset", on_click=reset_operator_filters, use_container_width=True)
            line_col, process_col, experience_col, zone_col = st.columns(4)
            with line_col:
                line = st.selectbox("Last line", ["All"] + sorted(view["Last line"].unique()), key="op_line")
            with process_col:
                process = st.selectbox("Process", ["All"] + sorted(view["Process"].astype(str).unique()), key="op_process")
            with experience_col:
                experience = st.selectbox("Experience", ["All"] + sorted(view["Experience"].astype(str).unique()), key="op_experience")
            with zone_col:
                zone = st.selectbox("Rate zone", ["All", "≥95%", "80–<95%", "60–<80%", "<60%", "Missing"], key="op_zone")
        filtered = view.copy()
        if search.strip():
            term = search.strip()
            mask = filtered["Operator"].astype(str).str.contains(term, case=False, regex=False) | filtered.index.astype(str).str.contains(term, case=False, regex=False)
            filtered = filtered.loc[mask]
        if process != "All":
            filtered = filtered.loc[filtered["Process"].eq(process)]
        if experience != "All":
            filtered = filtered.loc[filtered["Experience"].eq(experience)]
        if line != "All":
            filtered = filtered.loc[filtered["Last line"].eq(line)]
        rate = filtered["Rate"]
        masks = {"≥95%": rate.ge(95), "80–<95%": rate.ge(80) & rate.lt(95), "60–<80%": rate.ge(60) & rate.lt(80), "<60%": rate.lt(60), "Missing": rate.isna()}
        if zone != "All":
            filtered = filtered.loc[masks[zone]]
        filtered = filtered.sort_values("Rate", na_position="last", kind="stable")
        if mode == "Lowest 30":
            filtered = filtered.loc[filtered["Rate"].notna()].head(30)
        st.caption(f"{len(filtered)} of {len(view)} operators · Sorted by shift Rate · Filters use shift totals and last recorded line.")
        if filtered.empty:
            st.info("No operators match these filters.")
            return
        def colour_rate(value):
            if pd.isna(value):
                return "color:#788278"
            background, colour = ("#c6edcd", "#10582e") if value >= 95 else ("#ffe4a0", "#774500") if value >= 80 else ("#ffd2b5", "#963b0b") if value >= 60 else ("#ffc7c1", "#98251e")
            return f"background-color:{background};color:{colour};font-weight:bold"
        def request_operator_details(table_key, hour_lookup):
            selection = st.session_state.get(table_key, {}).get("selection", {})
            cells = selection.get("cells", [])
            if cells:
                position, column = cells[0]
                if 0 <= position < len(filtered):
                    st.session_state["op_pending_detail"] = (filtered.index[position], hour_lookup.get(column))

        @st.dialog("Operator details", width="large")
        def operator_dialog(operator_id, selected_hour=None):
            row = employee.loc[operator_id]
            info = view.loc[operator_id]
            st.markdown("""
            <style>
            .op-detail-title {font-size:23px;font-weight:750;color:#153d30;overflow-wrap:anywhere;}
            .op-detail-grid {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:16px 0;}
            .op-detail-card {background:#f0f5f1;border:1px solid #cfdfd3;border-radius:12px;padding:13px;min-width:0;}
            .op-detail-card span {font-size:12px;color:#526b5e;display:block;}
            .op-detail-card strong {font-size:23px;line-height:1.4;color:#123e30;display:block;overflow-wrap:anywhere;}
            .op-detail-card.hero {background:#176b50;}
            .op-detail-card.hero span,.op-detail-card.hero strong {color:white;}
            @media(max-width:640px) {
                .op-detail-grid {grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;}
                .op-detail-card {padding:10px;}
                .op-detail-card strong {font-size:20px;}
                .op-detail-title {font-size:20px;}
                div[role="dialog"] {max-width:calc(100vw - 16px)!important;max-height:92dvh;overflow-y:auto;}
            }
            </style>
            """, unsafe_allow_html=True)
            st.markdown(f'<div class="op-detail-title">{html.escape(str(info["Operator"]))}</div>',
                        unsafe_allow_html=True)
            st.caption(f"Employee ID: {operator_id} · Shift summary")
            def detail_cards(items):
                cards = []
                for label, value, hero in items:
                    cards.append(f'<div class="op-detail-card{" hero" if hero else ""}">'
                                 f'<span>{html.escape(label)}</span><strong>{html.escape(str(value))}</strong></div>')
                st.markdown('<div class="op-detail-grid">' + "".join(cards) + '</div>',
                            unsafe_allow_html=True)
            detail_cards([
                ("Shift Rate", number(info["Rate"], 2, "%"), True),
                ("Shift UPH", number(info["UPH"], 1), True),
                ("Experience", info["Experience"], False),
                ("Last station", info["Last station"], False),
                ("First scan", info["First scan"], False),
                ("Last scan", info["Last scan"], False),
                ("Worked hours", number(info["Hours"], 2), False),
                ("Break gap · min", number(info["Break time"], 1) if pd.notna(info["Break time"]) else "Not detected", False),
            ])
            if pd.isna(info["Break time"]):
                st.info(
                    "No qualifying break gap was detected for this operator (30–90 minutes overlapping a break slot)."
                    if break_column is not None else
                    "The employee report does not contain Total / Break gap min."
                )
                with st.expander("Break data diagnostics"):
                    st.write("Source column:", str(break_column) if break_column is not None else "No unambiguous duration column found")
                    st.write("Available break fields:",
                             [str(c) for c in employee.columns if "break" in str(c).lower()])
            else:
                st.caption("Break gap is the full scan-to-scan gap selected by find_break (30–90 minutes overlapping a break slot), not the scheduled 45-minute deduction.")
            choices = [None] + list(result["hours"])
            chosen = st.selectbox("Period", choices,
                index=choices.index(selected_hour) if selected_hour in choices else 0,
                format_func=lambda h: "Whole shift" if h is None else f"{(int(h)-1)%24:02d}:00–{int(h)%24:02d}:00",
                key=f"op_detail_period_{operator_id}_{selected_hour}")
            group = "Total" if chosen is None else chosen
            fields = [("Single", "Single"), ("Multi", "Multi"),
                      ("Units", "Quantity" if chosen is None else "Total"), ("Rate", "Rate")]
            period_cards = []
            for label, field in fields:
                value = pd.to_numeric(row.get((group, field), float("nan")), errors="coerce")
                period_cards.append((label, number(value, 2 if label == "Rate" else 0,
                                                   "%" if label == "Rate" else ""), False))
            detail_cards(period_cards)
            st.caption("Selected-period values come directly from the report. Single includes PM when they are combined in the report.")
            records = []
            for h in result["hours"]:
                for kind in ["Single", "Multi"]:
                    value = pd.to_numeric(row.get((h, kind), float("nan")), errors="coerce")
                    records.append({"Hour": f"{(int(h)-1)%24:02d}:00", "Type": kind, "Units": value})
            chart = alt.Chart(pd.DataFrame(records)).mark_bar().encode(
                x=alt.X("Hour:O", sort=list(dict.fromkeys(r["Hour"] for r in records))),
                y=alt.Y("Units:Q"), color=alt.Color("Type:N",
                    scale=alt.Scale(domain=["Single", "Multi"], range=["#e49a79", "#208463"])),
                tooltip=[alt.Tooltip("Hour:O"), alt.Tooltip("Type:N"),
                         alt.Tooltip("Units:Q", format=",.0f")]).properties(height=280)
            st.altair_chart(chart, use_container_width=True)
            with st.expander("Hourly values"):
                detail_rows = []
                for h in result["hours"]:
                    item = {"Hour": f"{(int(h)-1)%24:02d}:00–{int(h)%24:02d}:00"}
                    for label, field in [("Single", "Single"), ("Multi", "Multi"), ("Units", "Total"), ("Rate", "Rate"), ("UPH", "UPH")]:
                        if (h, field) in employee.columns:
                            item[label] = pd.to_numeric(row[(h, field)], errors="coerce")
                    detail_rows.append(item)
                st.dataframe(pd.DataFrame(detail_rows), hide_index=True, use_container_width=True)
            with st.expander("All report fields"):
                full = pd.DataFrame({
                    "Field": [" · ".join(map(str, c)) if isinstance(c, tuple) else str(c) for c in row.index],
                    "Value": ["—" if pd.isna(v) else str(v) for v in row.values],
                })
                st.dataframe(full, hide_index=True, use_container_width=True)

        st.caption("Select Open profile → or a name to view scan times, break and detailed results. Select an hourly cell for its Single / Multi breakdown.")
        table_token = hashlib.sha256(
            repr((list(filtered.index), display, st.session_state.get("packing_fingerprint"))).encode()
        ).hexdigest()[:12]
        if display == "Shift totals":
            shift_view = filtered[[
                "Operator", "Experience", "Last station", "Units", "Hours", "UPH", "Rate"
            ]].copy()
            shift_view["Profile"] = "Open profile →"
            styled = (shift_view.style
                      .set_properties(subset=["Units", "Hours", "UPH", "Rate"],
                                      **{"color": "#123e30", "font-weight": "bold"})
                      .map(colour_rate, subset=["Rate"])
                      .set_properties(subset=["Profile"], **{"color": "#176b50", "font-weight": "bold"}))
            with st.container(key="operator_shift_table"):
                st.dataframe(
                    styled, use_container_width=True, hide_index=True,
                    key=f"op_shift_{table_token}", selection_mode="single-cell",
                    on_select=lambda: request_operator_details(f"op_shift_{table_token}", {}),
                    height=min(650, 46 * (len(shift_view) + 1)), row_height=42,
                    column_config={
                        "Operator": st.column_config.TextColumn(width="medium", pinned=True),
                        "Last station": st.column_config.TextColumn(help="Last recorded station in the uploaded data."),
                        "Profile": st.column_config.TextColumn("Details", width="small", help="Click to open the operator profile."),
                        "Rate": st.column_config.NumberColumn(format="%.2f%%"),
                        "UPH": st.column_config.NumberColumn(format="%.1f"),
                        "Hours": st.column_config.NumberColumn(format="%.2f"),
                        "Units": st.column_config.NumberColumn(format="%.0f"),
                    },
                )
        else:
            # Only offer metrics that actually exist in the calculation report.
            hourly_fields = {"Rate": "Rate", "UPH": "UPH", "Units": "Total"}
            available = [label for label, field in hourly_fields.items()
                         if any((h, field) in employee.columns for h in result["hours"])]
            if not available:
                st.info("No hourly metrics are available in this report.")
            else:
                metric = st.radio("Hourly metric", available, horizontal=True, key="op_hourly_metric")
                hourly = filtered[["Operator", "Experience", "Last station"]].copy()
                hour_columns = []
                hour_lookup = {}
                for hour in result["hours"]:
                    column = (hour, hourly_fields[metric])
                    if column not in employee.columns:
                        continue
                    label = f"{(int(hour)-1)%24:02d}:00–{int(hour)%24:02d}:00"
                    hourly[label] = pd.to_numeric(employee.loc[filtered.index, column], errors="coerce")
                    hour_columns.append(label)
                    hour_lookup[label] = hour
                style = hourly.style.format(na_rep="—")
                if metric == "Rate":
                    style = style.map(colour_rate, subset=hour_columns)
                fmt = "%.2f%%" if metric == "Rate" else "%.1f" if metric == "UPH" else "%.0f"
                config = {label: st.column_config.NumberColumn(format=fmt) for label in hour_columns}
                config["Operator"] = st.column_config.TextColumn(width="medium", pinned=True)
                st.caption("Hours label the full interval. Missing values are shown as —; actual zeros remain zero. Ranking and filters still use shift totals.")
                hourly_key = f"op_hourly_{table_token}_{metric}"
                st.dataframe(style, use_container_width=True, hide_index=True,
                             key=hourly_key, selection_mode="single-cell",
                             on_select=lambda: request_operator_details(hourly_key, hour_lookup),
                             height=min(650, 46 * (len(hourly) + 1)), row_height=42,
                             column_config=config)
        pending = st.session_state.pop("op_pending_detail", None)
        if pending is not None and pending[0] in view.index:
            operator_dialog(*pending)
    render_people()

with lines:
    @st.fragment
    def render_lines_workspace():
        report = tables["line_report"].drop(index="Manual single", errors="ignore").copy()
        if report.empty:
            st.info("No line data available.")
            return
        report.index = report.index.map(str)
        shift_hours = list(result["hours"])
        label = lambda h: f"{(int(h)-1)%24:02d}:00"
        def value(line, hour, field):
            return pd.to_numeric(report.loc[line,(hour,field)],errors="coerce") if (hour,field) in report.columns else float("nan")
        totals = pd.DataFrame(index=report.index)
        for name, field in [("Units","Quantity"),("Rate","Rate"),("Single","Single"),("Multi","Multi")]:
            totals[name] = [value(line,"Total",field) for line in report.index]
        totals["Output/h"] = totals["Units"] / max(1,len(shift_hours))
        hourly_hc = tables.get("hc_final",pd.DataFrame())
        last_hc = tables.get("line_hc_report",pd.DataFrame())
        totals["HC"] = float("nan")
        for line in totals.index:
            if line in hourly_hc.index:
                sequence = pd.to_numeric(hourly_hc.loc[line],errors="coerce").dropna()
                if not sequence.empty:
                    totals.loc[line,"HC"] = sequence.iloc[-1]
            elif line in last_hc.columns and "Total" in last_hc.index:
                totals.loc[line,"HC"] = pd.to_numeric(last_hc.loc["Total",line],errors="coerce")
        st.markdown('''<style>
        .lccards{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:9px;margin:12px 0 22px}
        .lccard{background:white;border:1px solid #d6e5da;border-radius:13px;padding:13px;min-width:0}
        .lccard small{color:#60796b;display:block;font-size:12px}.lccard b{font-size:clamp(17px,2vw,27px);color:#176b50}
        @media(max-width:950px){.lccards{grid-template-columns:repeat(3,minmax(0,1fr))}}
        @media(max-width:650px){.lccards{grid-template-columns:repeat(2,minmax(0,1fr))}.lccard{padding:9px}}
        </style>''',unsafe_allow_html=True)
        def cards(items):
            st.markdown('<div class="lccards">'+''.join('<div class="lccard"><small>'+html.escape(k)+'</small><b>'+html.escape(number(v,d,s))+'</b></div>' for k,v,d,s in items)+'</div>',unsafe_allow_html=True)
        st.subheader("Line Control Center")
        st.caption("All lines at a glance · Select one to see hourly output and packing dynamics")
        cards([("Total output",totals["Units"].sum(),0,""),
               ("Avg total output/h",totals["Output/h"].sum(),0,""),
               ("Active lines",totals["Units"].gt(0).sum(),0,""),
               ("Lines below 95%",totals["Rate"].lt(95).sum(),0,""),
               ("Headcount (last hour)",pd.to_numeric(hourly_hc[shift_hours[-1]],errors="coerce").sum() if shift_hours[-1] in hourly_hc.columns else float("nan"),0,"")])
        left,right=st.columns([1,2.2],gap="medium")
        with left:
            st.markdown("#### Lines overview")
            search=st.text_input("Search line",placeholder="PB300009…",key="lcc_search")
            order=st.selectbox("Sort by",["Most units","Lowest Rate","Lowest units","Name"],key="lcc_sort")
            view=totals.loc[totals.index.str.contains(search.strip(),case=False,regex=False)].copy()
            if order=="Most units": view=view.sort_values("Units",ascending=False)
            elif order=="Lowest Rate": view=view.sort_values("Rate",na_position="last")
            elif order=="Lowest units": view=view.sort_values("Units")
            else: view=view.sort_index()
            if view.empty:
                st.info("No matching lines.")
                return
            if st.session_state.get("lcc_line") not in view.index: st.session_state["lcc_line"]=view.index[0]
            chosen=st.selectbox("Selected line",list(view.index),key="lcc_line")
            table_view=view[["Units","Output/h","HC","Rate"]]
            table_view.index.name="Line"
            st.dataframe(table_view.style.format({"Units":"{:,.0f}","Output/h":"{:,.0f}","HC":"{:.0f}","Rate":"{:.1f}%"},na_rep="—"),width="stretch",height=min(600,38*(len(view)+1)+10))
            st.caption("Output/h = units divided by reported hours; Rate is unchanged from source.")
        with right:
            st.markdown("### "+html.escape(chosen))
            cards([("Output",totals.loc[chosen,"Units"],0,""),
                   ("Avg output/h",totals.loc[chosen,"Output/h"],0,""),
                   ("Headcount",totals.loc[chosen,"HC"],0,""),
                   ("Rate",totals.loc[chosen,"Rate"],1,"%"),
                   ("Multi units",totals.loc[chosen,"Multi"],0,"")])
            hours_data=pd.DataFrame([{"Hour":label(h),"Single":value(chosen,h,"Single"),
                 "Multi":value(chosen,h,"Multi"),"Units":value(chosen,h,"Total"),
                 "Rate":value(chosen,h,"Rate"),"Key":h} for h in shift_hours])
            hour=st.selectbox("Inspect period",[None]+shift_hours,key="lcc_hour",
                 format_func=lambda h:"Whole shift" if h is None else label(h)+f"–{int(h)%24:02d}:00")
            a,b=st.columns(2,gap="small")
            with a:
                st.markdown("**Hourly output**")
                long=hours_data[["Hour","Single","Multi"]].melt(id_vars="Hour",var_name="Type",value_name="Units")
                chart=alt.Chart(long).mark_bar().encode(x=alt.X("Hour:O",sort=hours_data.Hour.tolist(),axis=alt.Axis(labelAngle=0)),y=alt.Y("Units:Q"),color=alt.Color("Type:N",scale=alt.Scale(domain=["Single","Multi"],range=["#e49a79","#176b50"])),tooltip=["Hour:N","Type:N",alt.Tooltip("Units:Q",format=",.0f")]).properties(height=260)
                st.altair_chart(chart,width="stretch")
            with b:
                st.markdown("**Packing dynamics · rolling 10 min × 6**")
                speed=tables.get("packing_speed_by_line")
                if speed is None or speed.empty:
                    st.info("No speed data.")
                else:
                    points=speed.loc[speed["Line"].astype(str).str.upper().str.strip().eq(chosen.upper())].copy()
                    points["Time"]=pd.to_datetime(points["Time"],errors="coerce")
                    points["Units per hour"]=pd.to_numeric(points["Units per hour"],errors="coerce")
                    points=points.dropna(subset=["Time","Units per hour"]).sort_values("Time")
                    if hour is not None:
                        points=points.loc[(points["Time"]-pd.Timedelta(nanoseconds=1)).dt.hour.eq((int(hour)-1)%24)]
                    if points.empty: st.info("No speed data for this period.")
                    else:
                        graph=alt.Chart(points).mark_area(color="#6aaf8d",opacity=.3,line={"color":"#176b50","strokeWidth":2}).encode(x=alt.X("Time:T",axis=alt.Axis(format="%H:%M",labelAngle=0)),y=alt.Y("Units per hour:Q",scale=alt.Scale(zero=True)),tooltip=[alt.Tooltip("Time:T",format="%H:%M"),alt.Tooltip("Units per hour:Q",format=",.0f")]).properties(height=260)
                        st.altair_chart(graph,width="stretch")
            st.markdown("**Hourly details**")
            display=hours_data if hour is None else hours_data.loc[hours_data.Key.eq(hour)]
            st.dataframe(display.drop(columns="Key").style.format({"Single":"{:,.0f}","Multi":"{:,.0f}","Units":"{:,.0f}","Rate":"{:.2f}%"},na_rep="—"),hide_index=True,width="stretch",height=min(410,37*(len(display)+1)+10))
            st.caption("Rate, hourly output and 10-minute speed are taken directly from existing calculations.")
        st.divider()
        st.subheader("Headcount by line · hourly")
        st.caption("Full hourly staffing table from the original report; no hours are hidden.")
        if not hourly_hc.empty:
            hc_display=hourly_hc.copy()
            hc_display.columns=[label(h) if isinstance(h,(int,float)) else str(h) for h in hc_display.columns]
            hc_display.index.name="Line / Process"
            st.dataframe(hc_display,width="stretch",height=min(620,38*(len(hc_display)+1)+10))
        else: st.info("No hourly headcount report available.")
        if not last_hc.empty:
            with st.expander("Headcount by experience · last available hour"):
                st.dataframe(last_hc.T,width="stretch")
        with st.expander("Full line report"):
            data_grid(report,"lcc_full_report")
    render_lines_workspace()

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