"""Streamlit UI. Усі обчислення виконує packing.calculations."""
import datetime as dt
import hashlib
import time
import pandas as pd
import streamlit as st
from packing.calculations import calculate
from packing.demo import demo_frames, as_excel
from packing.excel import export_excel
from packing.inputs import read_inputs, input_notes, PACK_COLUMNS, EMPLOYEE_COLUMNS

st.set_page_config(page_title="Pack Analytics · Звіт зміни", page_icon="📦", layout="wide")
st.markdown('''<style>
.block-container {padding-top:2.2rem; padding-bottom:3rem; max-width:1600px}
h1,h2,h3 {letter-spacing:-.035em}
[data-testid="stMetric"] {background:white;border:1px solid #E1E7EF;border-radius:14px;padding:18px 20px}
[data-testid="stMetricLabel"] {color:#64748B}
[data-testid="stMetricValue"] {font-weight:700}
.hero {background:linear-gradient(120deg,#112B40,#164D59);color:white;border-radius:20px;padding:30px 34px;margin-bottom:24px}
.hero h1 {color:white;margin:8px 0;font-size:2.5rem}
.hero p {color:#BFD8E0;margin-bottom:0;max-width:750px}
.eyebrow {font-size:.75rem;letter-spacing:.17em;text-transform:uppercase;color:#8FE0D3;font-weight:700}
[data-testid="stTabs"] [data-baseweb="tab-list"] {gap:18px}
[data-testid="stTabs"] [data-baseweb="tab"] {font-weight:600}
</style>''', unsafe_allow_html=True)
st.markdown('''<div class="hero"><div class="eyebrow">Operations intelligence / Test edition</div>
<h1>Pack Analytics</h1><p>Уся зміна в одному місці. Обсяги пакування, продуктивність команди та завантаження ліній.</p></div>''', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📦 Pack Analytics")
    st.caption("РОБОЧИЙ ПРОСТІР ЗМІНИ")
    mode = st.radio("Джерело даних", ["Демонстрація", "Мої Excel-файли"])
    pack_file = employee_file = None
    night = extended = False
    if mode == "Демонстрація":
        st.info("Вигадані дані · 12 операторів")
        night = st.toggle("Нічна зміна")
        extended = st.toggle("Продовжена зміна")
    else:
        pack_file = st.file_uploader("1. Операції пакування", type=["xlsx"], key="pack_upload")
        employee_file = st.file_uploader("2. Довідник працівників", type=["xlsx"], key="employees_upload")
        st.caption("Наприклад: Pack_data.xlsx та InfUser.xlsx. Читається перший аркуш кожного файлу.")
    run = st.button("Сформувати звіт", type="primary", use_container_width=True)
    st.divider()
    st.caption("Логіка оригіналу")
    st.markdown("**120 / 260 / 315** — базові норми\n\n**>15** — поріг активної години\n\n**>100** — поріг звіту працівників")
    st.caption("Стаж обчислюється на поточну дату сервера, як у вихідному коді.")

if mode == "Мої Excel-файли" and (pack_file is None or employee_file is None):
    st.info("Завантажте обидва Excel-файли зліва та натисніть «Сформувати звіт».")
    with st.expander("Які колонки потрібні?", expanded=True):
        c1, c2 = st.columns(2)
        c1.markdown("**Операції пакування**\n\n" + "\n".join(f"- `{c}`" for c in PACK_COLUMNS))
        c2.markdown("**Довідник працівників**\n\n" + "\n".join(f"- `{c}`" for c in EMPLOYEE_COLUMNS))
        st.caption("Operator має містити ID у дужках, наприклад Оператор (1001). Завантажуйте дані однієї зміни.")
    st.stop()

if mode == "Демонстрація":
    fingerprint = f"demo-{night}-{extended}-{dt.date.today()}"
else:
    pack_bytes, employee_bytes = pack_file.getvalue(), employee_file.getvalue()
    fingerprint = hashlib.sha256(pack_bytes + employee_bytes + str(dt.date.today()).encode()).hexdigest()

if run or (mode == "Демонстрація" and st.session_state.get("fingerprint") != fingerprint):
    st.session_state.pop("result", None)
    st.session_state.pop("fingerprint", None)
    try:
        with st.spinner("Обчислюємо показники та готуємо Excel…"):
            started = time.perf_counter()
            pack, employees = demo_frames(night, extended) if mode == "Демонстрація" else read_inputs(pack_bytes, employee_bytes)
            notes = input_notes(pack, employees)
            result = calculate(pack, employees)
            excel = export_excel(result["tables"])
            st.session_state.update(result=result, excel=excel, notes=notes,
                                    elapsed=time.perf_counter()-started, fingerprint=fingerprint)
    except Exception as exc:
        st.error("Не вдалося сформувати звіт для цих даних. Перевірте формат колонок і склад зміни.")
        st.caption("Оригінальний алгоритм очікує наявність Single і Multi та придатних даних для погодинних звітів. Формули автоматично не змінюються.")
        with st.expander("Деталі помилки"):
            st.code(f"{type(exc).__name__}: {exc}")
        st.stop()

if st.session_state.get("fingerprint") != fingerprint or "result" not in st.session_state:
    st.info("Файли готові. Натисніть «Сформувати звіт», щоб розрахувати нові дані.")
    st.stop()

result = st.session_state.result
tables = result["tables"]
for note in st.session_state.notes:
    st.warning(note)

left, right = st.columns([3, 1])
with left:
    shift = "Нічна" if result["night"] else "Денна"
    st.subheader(f"Звіт зміни · {shift}")
    st.caption(f"{'ДЕМОНСТРАЦІЯ · ' if mode == 'Демонстрація' else ''}{'Продовжена зміна · ' if result['extended'] else ''}Дата розрахунку: {result['calculated_on']} · {len(result['hours'])} годин у звіті · {st.session_state.elapsed:.1f} с")
with right:
    st.download_button("↓ Завантажити Excel", st.session_state.excel, "Pack_report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", use_container_width=True)

stats = tables["statistics_report"].set_index("Metric")["Value"]
cols = st.columns(4)
for col, label, key in zip(cols, ["Усього одиниць", "Single", "Multi", "Операторів"],
                           ["Total Quantity", "Total Single", "Total Multi", "Total Operators"]):
    value = stats[key]
    col.metric(label, "—" if pd.isna(value) else f"{value:,.0f}".replace(",", " "))
st.write("")


def show_table(title, key, caption=None):
    st.markdown(f"#### {title}")
    if caption:
        st.caption(caption)
    # Use a display copy: original values and MultiIndex remain intact in Excel.
    view = tables[key].copy()
    if isinstance(view.columns, pd.MultiIndex):
        view.columns = [" · ".join(str(part) for part in column if str(part)) for column in view.columns]
    else:
        view.columns = view.columns.map(str)
    if key == "Libiao":
        view = view.astype(str)
    st.dataframe(view, use_container_width=True, height=min(600, max(180, 36 * (len(view) + 1))),
                 hide_index=key in ["statistics_report", "rate_zone_report"])


overview, operators, lines, statistics, details = st.tabs(["Огляд зміни", "Працівники", "Лінії", "Статистика", "Про звіт"])
with overview:
    st.markdown("#### Динаміка пакування")
    employee = tables["employee_report"]
    # The line report contains a duplicate summary row; use only actual lines.
    actual_lines = tables["line_report"].drop(index="Manual single", errors="ignore")
    hourly = pd.DataFrame({kind: [actual_lines[(hour, kind)].sum() for hour in result["hours"]]
                           for kind in ["Single", "Multi"]},
                          index=[f"{hour-1:02}:00–{hour%24:02}:00" for hour in result["hours"]])
    chart_data = hourly.rename_axis("Година").reset_index().melt(
        id_vars="Година", var_name="Тип", value_name="Кількість")
    st.vega_lite_chart(chart_data, {
        "height": 260,
        "mark": {"type": "bar", "cornerRadiusTopLeft": 3, "cornerRadiusTopRight": 3},
        "encoding": {
            "x": {"field": "Година", "type": "ordinal", "sort": list(hourly.index),
                  "axis": {"labelAngle": -35}},
            "y": {"field": "Кількість", "type": "quantitative"},
            "color": {"field": "Тип", "type": "nominal",
                      "scale": {"domain": ["Single", "Multi"], "range": ["#087F8C", "#A6D8CF"]}},
            "tooltip": [{"field": "Година"}, {"field": "Тип"}, {"field": "Кількість"}]
        }
    }, use_container_width=True)
    st.caption("Години та обсяги взяті з готового звіту. Правила перенесення операцій о 05:55 / 17:55 збережено.")
    show_table("Dashboard · HC, UPH та Rate", "dashboard_report", "Ті самі значення, що на аркуші Dashboard.")
    a, b = st.columns([1, 2])
    with a:
        show_table("Зони Rate", "rate_zone_report", "shape — частка операторів у %, назву збережено з оригіналу.")
    with b:
        show_table("Погодинний HC", "hc_final", "Кількість активних працівників за напрямком роботи.")
with operators:
    st.caption("ByOperator · працівники з понад 100 одиницями за години звіту. Пошук і сортування доступні в таблиці.")
    show_table("Продуктивність працівників", "employee_report")
    with st.expander("Додатково: деталізація по станціях"):
        show_table("Станції", "station_report_display", "Готова проміжна таблиця оригінального коду; не входить до Excel.")
with lines:
    show_table("Погодинний звіт по лініях", "line_report", "By Line · включно з підсумковим рядком Manual single.")
    show_table("Поточний HC по лініях", "line_hc_report", "Остання доступна година звіту, як у вихідному алгоритмі.")
    show_table("Відповідність Libiao", "Libiao")
with statistics:
    show_table("Усі показники зміни", "statistics_report", "Повна таблиця з аркуша Statistics.")
with details:
    st.markdown("""#### Як працює звіт
1. Завантажте операції та довідник працівників у форматі `.xlsx`.
2. Натисніть **Сформувати звіт**. Зміна та її продовження визначаються оригінальним алгоритмом.
3. Перегляньте таблиці та завантажте готовий **Pack_report.xlsx**.

**Збережено:** норми, коефіцієнти стажу, пошук перерви, пороги, правила HC, UPH, Rate та порядок розрахунків.

Усі 8 таблиць Excel відображаються на сайті. Позначка `—` означає відсутнє значення, а не нуль. На графіку не дублюється підсумковий рядок Manual single.

Файли обробляються в пам’яті поточної сесії. Після перезапуску сервера звіт потрібно сформувати знову.
""")
    if mode == "Демонстрація":
        d1, d2 = st.columns(2)
        demo_pack, demo_employees = demo_frames(night, extended)
        d1.download_button("Приклад операцій .xlsx", as_excel(demo_pack), "Demo_Pack.xlsx")
        d2.download_button("Приклад довідника .xlsx", as_excel(demo_employees), "Demo_Employees.xlsx")
st.caption("PACK ANALYTICS / TEST EDITION · Розрахункове ядро відокремлене від інтерфейсу")
