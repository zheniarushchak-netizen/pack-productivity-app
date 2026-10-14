# ======================================================================
# 01. ІМПОРТИ
# ======================================================================
# Підключаємо бібліотеки для таблиць і роботи з датами.

import pandas as pd

import datetime as dt
from time import perf_counter
from openpyxl.styles import PatternFill, Font, Alignment

# ======================================================================
# 01. STYLE EXCEL
# ======================================================================

def write_table(
    writer,
    dataframe,
    sheet_name,
    startrow=0,
    startcol=0,
    index=True,
    header=True,
    format_headers=True
):
    dataframe.to_excel(
        writer,
        sheet_name=sheet_name,
        startrow=startrow,
        startcol=startcol,
        index=index,
        header=header
    )

    if format_headers:
        worksheet = writer.book[sheet_name]

        format_table_headers(
            worksheet,
            dataframe,
            startrow=startrow,
            startcol=startcol,
            index=index
        )
    # ======================================================================

def format_table_headers(
    worksheet,
    dataframe,
    startrow=0,
    startcol=0,
    index=True,
    fill_color="D9EAF7"
):
    header_fill = PatternFill(
        fill_type="solid",
        fgColor=fill_color
    )

    header_font = Font(
        bold=True
    )

    header_alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    excel_start_row = startrow + 1
    excel_start_col = startcol + 1

    header_rows = dataframe.columns.nlevels

    if index:
        index_cols = dataframe.index.nlevels
    else:
        index_cols = 0

    table_width = len(dataframe.columns) + index_cols

    # -------------------------
    # Заголовки колонок
    # -------------------------
    for row in range(
        excel_start_row,
        excel_start_row + header_rows
    ):
        for col in range(
            excel_start_col,
            excel_start_col + table_width
        ):
            cell = worksheet.cell(
                row=row,
                column=col
            )

            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

    # MultiIndex columns створюють
    # додатковий рядок для назви index
    extra_row = 0

    if index and header_rows > 1:
        extra_row = 1

    first_data_row = (
        excel_start_row
        + header_rows
        + extra_row
    )

    # -------------------------
    # Index таблиці
    # -------------------------
    if index:
        for row in range(
            first_data_row,
            first_data_row + len(dataframe.index)
        ):
            for col in range(
                excel_start_col,
                excel_start_col + index_cols
            ):
                cell = worksheet.cell(
                    row=row,
                    column=col
                )

                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
# ======================================================================
# 02. ДОПОМІЖНІ ФУНКЦІЇ
# ======================================================================
# Функції визначають стаж, перерву та Rate. get_status використовує today, задану нижче.

def get_status(create_time):
    work_days = []
    exp_status=[]
    for start_data in create_time:
        exp_days = today - start_data
        exp_days = exp_days.days
        work_days.append(exp_days)
        if exp_days == 0:
            exp_status.append("Day1")
        elif exp_days <=7:
            exp_status.append("Week 1")
        elif exp_days <=14:
            exp_status.append("Week 2")
        elif exp_days <=21:
            exp_status.append("Week 3")
        else:
            exp_status.append("Veteran")
    return exp_status, work_days

def find_break(
    operation_times,
    is_night_shift,
    min_gap_minutes=30,
    max_gap_minutes=90
):

    times = (
        operation_times
        .dropna()
        .sort_values()
        .drop_duplicates()
        .reset_index(drop=True)
    )
    if is_night_shift:

        default_distribution = {
            23: 0.75
        }

        default_break_name = "22:00-22:45"

    else:

        default_distribution = {
            11: 0.75
        }

        default_break_name = "10:00-10:45"
    # Якщо маємо менше двох операцій,
    # знайти перерву неможливо
    if len(times) < 2:
        return (
    default_distribution,
    default_break_name,
    0
)


    # Різниця між кожною операцією і попередньою
    gaps = (
        times
        .diff()
        .dt.total_seconds()
        / 60
    )

    # Усі проміжки, які можуть відповідати перерві
    candidate_gaps = gaps.loc[
        (gaps >= min_gap_minutes)&
        (gaps<=max_gap_minutes)
    ]

      # Якщо кандидатів немає, використовуємо стандартну перерву
    if candidate_gaps.empty:
        return (
            default_distribution,
            default_break_name,
            0
        )

    # Початкову дату використовуємо для створення слотів перерви
    day = times.iloc[0].normalize()


    if is_night_shift:

        break_options = [

        {
            "name": "21:45-22:30",
            "start": day + pd.Timedelta(
                hours=21,
                minutes=45
            ),
            "end": day + pd.Timedelta(
                hours=22,
                minutes=30
            ),
            "distribution": {
                22: 0.25,
                23: 0.50
            }
        },

        {
            "name": "22:00-22:45",
            "start": day + pd.Timedelta(hours=22),
            "end": day + pd.Timedelta(
                hours=22,
                minutes=45
            ),
            "distribution": {
                23: 0.75
            }
        },

        {
            "name": "22:15-23:00",
            "start": day + pd.Timedelta(
                hours=22,
                minutes=15
            ),
            "end": day + pd.Timedelta(hours=23),
            "distribution": {
                23: 0.75
            }
        },

        {
            "name": "22:30-23:15",
            "start": day + pd.Timedelta(
                hours=22,
                minutes=30
            ),
            "end": day + pd.Timedelta(
                hours=23,
                minutes=15
            ),
            "distribution": {
                23: 0.50,
                24: 0.25
            }
        }
    ]

    else:

        break_options = [

        {
            "name": "09:45-10:30",
            "start": day + pd.Timedelta(
                hours=9,
                minutes=45
            ),
            "end": day + pd.Timedelta(
                hours=10,
                minutes=30
            ),
            "distribution": {
                10: 0.25,
                11: 0.50
            }
        },

        {
            "name": "10:00-10:45",
            "start": day + pd.Timedelta(hours=10),
            "end": day + pd.Timedelta(
                hours=10,
                minutes=45
            ),
            "distribution": {
                11: 0.75
            }
        },

        {
            "name": "10:15-11:00",
            "start": day + pd.Timedelta(
                hours=10,
                minutes=15
            ),
            "end": day + pd.Timedelta(hours=11),
            "distribution": {
                11: 0.75
            }
        },

        {
            "name": "10:30-11:15",
            "start": day + pd.Timedelta(
                hours=10,
                minutes=30
            ),
            "end": day + pd.Timedelta(
                hours=11,
                minutes=15
            ),
            "distribution": {
                11: 0.50,
                12: 0.25
            }
        }
    ]



    best_break = None
    best_overlap = 0
    best_gap_minutes = 0

    # Перебираємо всі проміжки від min_gap_minutes
    for position, gap_minutes in candidate_gaps.items():

        gap_start = times.iloc[position - 1]
        gap_end = times.iloc[position]

        # Порівнюємо проміжок із кожним слотом перерви
        for option in break_options:

            # Переносимо слот на дату початку цього проміжку
            date_offset = gap_start.normalize() - day

            option_start = option["start"] + date_offset
            option_end = option["end"] + date_offset

            overlap_start = max(gap_start, option_start)
            overlap_end = min(gap_end, option_end)

            overlap_minutes = max(
                0,
                (overlap_end - overlap_start).total_seconds() / 60
            )

            # Зберігаємо пару з найбільшим перетином
            if overlap_minutes > best_overlap:
                best_overlap = overlap_minutes
                best_break = option
                best_gap_minutes = gap_minutes

    # Жоден кандидат не перетнувся зі слотами
    if best_break is None:
        return (
            default_distribution,
            default_break_name,
            0
        )

    return (
        best_break["distribution"],
        best_break["name"],
        best_gap_minutes
    )
def cal_rate(
    hours,
    work_level,
    break_distribution=None,
    last_shift_hour=None,
    base_single_norma=120,
    base_multi_norma=260,
    min_quantity=15
):
    if break_distribution is None:
        break_distribution = {}

    work_levels = {
        "Day1":0.55,
        "Week 1": 0.65,
        "Week 2": 0.70,
        "Week 3": 0.85,
        "Veteran": 1.00,
        "Unknown": 1.00
    }

    level_rate = work_levels[work_level]

    single_norma = base_single_norma * level_rate
    multi_norma = base_multi_norma * level_rate

    single_total_hours = 0
    multi_total_hours = 0

    valid_hours = 0
    ignored_hours = 0
    worked_hours = 0

    hourly_rates = {}


    for hour in hours:

        single_quantity, multi_quantity = hours[hour]

        total_quantity = (
            single_quantity
            + multi_quantity
        )

        # Якщо за годину зроблено 15 або менше,
        # цю годину не враховуємо в Rate
        if total_quantity <= min_quantity:

            hourly_rates[hour] = pd.NA
            ignored_hours += 1

            break_time = break_distribution.get(
                hour,
            0
            )

            # Якщо це година перерви
            # і працівник зробив більше 5 pack
            if (
                break_time > 0
                and total_quantity > 5
                ):
                worked_hours += (
                    1 - break_time
                    )

            continue
        # Скільки "нормо-годин" виконано
        # окремо для Single та Multi
        single_norma_hour = (
            single_quantity
            / single_norma
        )

        multi_norma_hour = (
            multi_quantity
            / multi_norma
        )

        single_total_hours += (
            single_norma_hour
        )

        multi_total_hours += (
            multi_norma_hour
        )

        # Rate конкретної години
        break_time = break_distribution.get(
            hour,
            0
        )

                # Останні 15 хвилин зміни не працюємо
        if hour == last_shift_hour:
            end_shift_time = 0.25
        else:
            end_shift_time = 0

        hour_worked_time = (
            1 - break_time - end_shift_time
        )

        hour_rate = round(
            (
                single_norma_hour
                + multi_norma_hour
            )
            / hour_worked_time
            * 100,
            2
        )

        hourly_rates[hour] = hour_rate

        worked_hours += hour_worked_time

        valid_hours += 1


    total_norma_hours = (
        single_total_hours
        + multi_total_hours
    )
    

    # Загальний середній Rate
    if worked_hours > 0:
 
        total_avg = round(
            total_norma_hours
            / worked_hours
            * 100,
            2
        )

    else:

        total_avg = 0


    return (
    single_total_hours,
    multi_total_hours,
    total_norma_hours,
    valid_hours,
    ignored_hours,
    worked_hours,
    total_avg,
    hourly_rates
)

# ======================================================================
# 03. ВХІДНІ КОЛОНКИ ТА ЧИТАННЯ ФАЙЛІВ
# ======================================================================
# Читаємо лише потрібні колонки з двох вхідних Excel-файлів.

pack_columns = [
    "Package No.",
    "Item quantity classification name",
    "Operator",
    "Packing station",
    "Operation time",
    "Quantity of scanned"
]

employee_columns = [
    "Employee ID",
    "Create time"]

data = pd.read_excel(
    "Pack_data_06.10.2026.xlsx",
    usecols=pack_columns,
    engine="openpyxl"
)

data_employee = pd.read_excel(
    "InfUser.xlsx",
    usecols=employee_columns,
    engine="openpyxl"
    )



# ======================================================================
# 04. ОЧИЩЕННЯ ТА ПІДГОТОВКА ДАНИХ
# ======================================================================
# Приводимо дати, ID і кількість до потрібних типів; відкидаємо рядки без основних даних.

data_employee["Create time"]=(pd.to_datetime(data_employee["Create time"]).dt.normalize())

today =pd.to_datetime(dt.date.today())

create_time = data_employee["Create time"]

data_employee["Employee ID"] = (
    data_employee["Employee ID"]
    .astype("Int64")
    .astype("string")
)

data["Operation time"] = pd.to_datetime(
    data["Operation time"],
    errors="coerce"
)

data["Quantity of scanned"] = pd.to_numeric(
    data["Quantity of scanned"],
    errors="coerce"
).fillna(0)

data = data.dropna(
    subset=[
        "Operator",
        "Packing station",
        "Operation time"
    ]
)

data["Packing station"] = (
    data["Packing station"]
    .astype(str)
    .str.upper()
)



# ======================================================================
# 05. ПОХІДНІ КОЛОНКИ: EMPLOYEE ID
# ======================================================================
# Витягуємо ID з дужок у назві оператора. Якщо ID відсутній, залишається пропущене значення.

data["Employee ID"] = (
    data["Operator"]
    .astype(str)
    .str.extract(r"\((\d+)\)")[0]
)

data["Employee ID"] = (
    data["Employee ID"]
    .astype("string")
)



# ======================================================================
# 06. ПОХІДНІ КОЛОНКИ: PACK TYPE
# ======================================================================
# Визначаємо Single / Multi та залишаємо початковий діагностичний вивід.

data["Pack Type"] = (
    data[
        "Item quantity classification name"
    ]
    .apply(
        lambda x:
        "Single"
        if x == "Single product"
        else "Multi"
    )
)

print(
    "BEFORE MERGE:",
    data.loc[
        data["Pack Type"] == "Single",
        "Quantity of scanned"
    ].sum()
)



# ======================================================================
# 07. ПОХІДНІ КОЛОНКИ: HOUR І LINE
# ======================================================================
# Створюємо номер години та лінію, зберігаючи наявні правила для Pack Machine.

data["Hour"] = (
    data["Operation time"]
    .dt.hour
    + 1
)

data.loc[
    (data["Operation time"].dt.hour == 5)
    & (data["Operation time"].dt.minute >= 55),
    "Hour"
] = 7

data.loc[
    (data["Operation time"].dt.hour == 17)
    & (data["Operation time"].dt.minute >= 55),
    "Hour"
] = 19

data["Line"] = (
    data["Packing station"]
    .astype(str)
    .str.split("-")
    .str[0]
)

data.loc[ 
    data["Packing station"]
    .astype(str)
    .str.contains("PLA")|
    data["Packing station"]
    .astype(str)
    .str.contains("PL00"),
    "Line"
] = "Pack Machine"

data.loc[ 
    data["Packing station"]
    .astype(str)
    .str.contains("PLE"),
    "Line"
] = "Line 0"



# ======================================================================
# 08. ПОХІДНІ КОЛОНКИ: WORK LEVEL
# ======================================================================
# Розраховуємо стаж, приєднуємо його за Employee ID та заповнюємо відсутні значення.

exp_status, work_days = get_status(create_time)

data_employee["Work level"] = exp_status

data_employee["Experience days"] = work_days

data = data.merge(
    data_employee[
        [
            "Employee ID",
            "Work level",
            "Experience days"
        ]
    ],
    on="Employee ID",
    how="left"
)

data["Work level"] = (
    data["Work level"]
    .fillna("Unknown")
)

data["Experience days"] = (
    data["Experience days"]
    .fillna(-1)
)



# ======================================================================
# 09. НАЛАШТУВАННЯ ГОДИН ЗВІТУ
# ======================================================================
# Визначаємо Day / Night і порядок доступних годин за наявними умовами.

is_night_shift = (
    (
        data["Operation time"].dt.hour >= 18
    )
    |
    (
        (data["Operation time"].dt.hour == 17)
        & (data["Operation time"].dt.minute >= 55)
    )
).any()
#------------------------
#Маска яка фіксує проміжок
EXTENSION_THRESHOLD = 200

if is_night_shift:
    extension_mask = (
        (data["Operation time"].dt.hour == 4)
        & (data["Operation time"].dt.minute >= 50)
    )
else:
    extension_mask = (
        (data["Operation time"].dt.hour == 16)
        & (data["Operation time"].dt.minute >= 50)
    )

extension_quantity = data.loc[
    extension_mask,
    "Quantity of scanned"
].sum()

extended_shift = extension_quantity > EXTENSION_THRESHOLD

if is_night_shift:
    if extended_shift:
        last_shift_hour = 6
    else:
        last_shift_hour = 5
else:
    if extended_shift:
        last_shift_hour = 18
    else:
        last_shift_hour = 17
#============================================================
if is_night_shift:

    hours_order = [
        19, 20, 21, 22, 23, 24,
        1, 2, 3, 4, 5, 6
    ]

else:

    hours_order = [
        7, 8, 9, 10, 11, 12,
        13, 14, 15, 16, 17, 18
    ]

available_hours = set(
    data["Hour"]
    .dropna()
    .astype(int)
    .unique()
)

hours_list = [
    hour
    for hour in hours_order
    if hour in available_hours
]




# ==========================================================
# ЗВІТ ПО ПРАЦІВНИКАХ: ПІДГОТОВКА ДАНИХ
# ==========================================================

employee_report_start = perf_counter()

employee_grouped = (
    data
    .groupby(
        [
            "Employee ID",
            "Hour",
            "Pack Type"
        ],
        dropna=False
    )["Quantity of scanned"]
    .sum()
    .reset_index()
)
# Погодинна таблиця: один рядок на працівника
employee_hourly = (
    employee_grouped
    .pivot_table(
        index="Employee ID",
        columns=["Hour", "Pack Type"],
        values="Quantity of scanned",
        aggfunc="sum",
        fill_value=0
    )
)
for hour in hours_list:
    for pack_type in ["Single", "Multi"]:
        if (hour, pack_type) not in employee_hourly.columns:
            employee_hourly[(hour, pack_type)] = 0

    employee_hourly[(hour, "Total")] = (
        employee_hourly[(hour, "Single")]
        + employee_hourly[(hour, "Multi")]
    )

employee_hourly[("Total", "Single")] = (
    employee_hourly[
        [(hour, "Single") for hour in hours_list]
    ]
    .sum(axis=1)
)

employee_hourly[("Total", "Multi")] = (
    employee_hourly[
        [(hour, "Multi") for hour in hours_list]
    ]
    .sum(axis=1)
)

# Час операцій, згрупований за працівником
employee_operation_groups = (
    data.groupby("Employee ID")["Operation time"]
)

# Підсумкова кількість юнітів за години звіту
employee_hourly[("Total", "Quantity")] = (
    employee_hourly[
        [(hour, "Total") for hour in hours_list]
    ]
    .sum(axis=1)
)

# Колонки для наступних розрахунків
employee_hourly[("Total", "Worked hours")] = 0.0
employee_hourly[("Total", "UPH")] = 0.0
employee_hourly[("Total", "Break slot")] = ""
employee_hourly[("Total", "Break gap min")] = 0.0

for hour in hours_list:
    employee_hourly[(hour, "Worked hours")] = 0.0

for employee_id, row in employee_hourly.iterrows():

    # Усі операції працівника з усіх станцій
    employee_operations = (
        employee_operation_groups.get_group(employee_id)
    )

    break_distribution, break_slot, break_gap_minutes = find_break(
        employee_operations,
        is_night_shift
    )

    employee_hourly.loc[
        employee_id,
        ("Total", "Break slot")
    ] = break_slot

    employee_hourly.loc[
        employee_id,
        ("Total", "Break gap min")
    ] = round(break_gap_minutes, 2)

    worked_hours = 0.0

    for hour in hours_list:
        quantity = row[(hour, "Total")]

        break_time = break_distribution.get(hour, 0)

        if hour == last_shift_hour:
            end_shift_time = 0.25
        else:
            end_shift_time = 0

        if quantity <= 15:
            if break_time > 0 and quantity > 5:
                hour_worked_time = 1 - break_time
            else:
                hour_worked_time = 0.0
        else:
            hour_worked_time = (
                1 - break_time - end_shift_time
            )

        employee_hourly.loc[
            employee_id,
            (hour, "Worked hours")
        ] = hour_worked_time

        worked_hours += hour_worked_time

    employee_hourly.loc[
        employee_id,
        ("Total", "Worked hours")
    ] = worked_hours



# Ділимо тільки там, де є врахований робочий час
worked_mask = (
    employee_hourly[("Total", "Worked hours")] > 0
)

# Без робочого часу UPH залишається невизначеним
employee_hourly[("Total", "UPH")] = float("nan")

employee_hourly.loc[
    worked_mask,
    ("Total", "UPH")
] = (
    employee_hourly.loc[worked_mask, ("Total", "Quantity")]
    / employee_hourly.loc[worked_mask, ("Total", "Worked hours")]
).round(2)

# ======================================================================
# 10. ГРУПУВАННЯ ТА ЗВЕДЕНА ТАБЛИЦЯ СТАНЦІЙ
# ======================================================================
# Підсумовуємо кількість і розгортаємо Single / Multi в колонки за годинами.
# Ім’я та рівень досвіду кожного працівника
employee_details = (
    data
    .groupby("Employee ID")[["Operator", "Work level"]]
    .first()
)

employee_hourly[("Operator", "")] = employee_details["Operator"]
employee_hourly[("Work level", "")] = employee_details["Work level"]

# Остання станція за часом операції в годинах звіту
employee_last_station = (
    data.loc[data["Hour"].isin(hours_list)]
    .sort_values("Operation time")
    .groupby("Employee ID")["Packing station"]
    .last()
)

employee_hourly[("Last station", "")] = employee_last_station



station_grouped = (
    data
    .groupby(
        [
            "Employee ID",
            "Operator",
            "Packing station",
            "Hour",
            "Pack Type",
            "Work level"
        ],
        dropna=False
    )["Quantity of scanned"]
    .sum()
    .reset_index()
)

station_hourly = (
    station_grouped
    .pivot_table(
        index=[
            "Employee ID",
            "Operator",
            "Packing station",
            "Work level"
        ],
        columns=[
            "Hour",
            "Pack Type"
        ],
        values="Quantity of scanned",
        aggfunc="sum",
        fill_value=0
    )
)



# ======================================================================
# 11. ПОГОДИННІ ТА ПІДСУМКОВІ КОЛОНКИ
# ======================================================================
# Додаємо відсутні Single / Multi, готуємо числові Rate та підсумкові колонки.

for hour in hours_list:

    if (hour, "Single") not in station_hourly.columns:

        station_hourly[
            (hour, "Single")
        ] = 0


    if (hour, "Multi") not in station_hourly.columns:

        station_hourly[
            (hour, "Multi")
        ] = 0

    # Загальна кількість юнітів за кожну годину
    for hour in hours_list:
        station_hourly[(hour, "Total")] = (
            station_hourly[(hour, "Single")]
            + station_hourly[(hour, "Multi")]
        )
for hour in hours_list:

    station_hourly[
        (hour, "Rate")
    ] = pd.NA

for hour in hours_list:

    station_hourly[
        (hour, "Rate")
    ] = (
        station_hourly[
            (hour, "Rate")
        ]
        .astype("Float64")
    )

station_hourly[
    ("Total", "Single")
] = 0.0

station_hourly[
    ("Total", "Multi")
] = 0.0

station_hourly[
    ("Total", "Quantity")
] = 0.0

station_hourly[
    ("Total", "Valid hours")
] = 0

station_hourly[
    ("Total", "Ignored hours")
] = 0

station_hourly[
    ("Total", "Worked hours")
] = 0.0

station_hourly[
    ("Total", "Break slot")
] = ""

station_hourly[
    ("Total", "Break gap min")
] = 0.0

station_hourly[
    ("Total", "Rate")
] = 0.0



# ======================================================================
# 12. CAL_RATE ДЛЯ ПРАЦІВНИКА НА СТАНЦІЇ
# ======================================================================
# Для кожного рядка визначаємо перерву, передаємо погодинні кількості у cal_rate та записуємо результати.

for index, row in station_hourly.iterrows():

    employee_id = index[0]
    operator = index[1]
    station = index[2]
    work_level = index[3]

    employee_operations = data.loc[
        data["Employee ID"] == employee_id,
        "Operation time"
    ]

    (
        break_distribution,
        break_slot,
        break_gap_minutes

    )= find_break(
    employee_operations,
    is_night_shift
)

    hours = {}

    for hour in hours_list:

        single_quantity = row.get(
            (hour, "Single"),
            0
        )

        multi_quantity = row.get(
            (hour, "Multi"),
            0
        )

        if pd.isna(single_quantity):
            single_quantity = 0

        if pd.isna(multi_quantity):
            multi_quantity = 0

        hours[hour] = (
            float(single_quantity),
            float(multi_quantity)
        )

        if station.startswith(("PLA", "PL0")):
            current_single_norma = 315
        else:
            current_single_norma = 120
    (
        single_total_hours,
        multi_total_hours,
        total_norma_hours,
        valid_hours,
        ignored_hours,
        worked_hours,
        total_avg,
        hourly_rates

      ) = cal_rate(
        hours,
        work_level,
        break_distribution,
        last_shift_hour=last_shift_hour,
        base_single_norma=current_single_norma
    )


    for hour, rate in hourly_rates.items():

        station_hourly.loc[
            index,
            (hour, "Rate")
        ] = rate



    total_single = sum(
        hours[hour][0]
        for hour in hours
    )

    total_multi = sum(
        hours[hour][1]
        for hour in hours
    )


    total_quantity = (
        total_single
        + total_multi
    )


    station_hourly.loc[
        index,
        ("Total", "Single")
    ] = total_single


    station_hourly.loc[
        index,
        ("Total", "Multi")
    ] = total_multi


    station_hourly.loc[
        index,
        ("Total", "Quantity")
    ] = total_quantity


    station_hourly.loc[
        index,
        ("Total", "Valid hours")
    ] = valid_hours


    station_hourly.loc[
        index,
        ("Total", "Ignored hours")
    ] = ignored_hours
    station_hourly.loc[
        index,
        ("Total", "Worked hours")
    ] = worked_hours

    station_hourly.loc[
    index,
    ("Total", "Single norm hours")
    ] = single_total_hours

    

    station_hourly.loc[
        index,
        ("Total", "Break slot")
    ] = break_slot


    station_hourly.loc[
        index,
        ("Total", "Break gap min")
    ] = round(
        break_gap_minutes,
        1
    )

    station_hourly.loc[
        index,
        ("Total", "Rate")
    ] = total_avg



# ======================================================================
# 13. ОФОРМЛЕННЯ ЗВІТУ ПО СТАНЦІЯХ
# ======================================================================
# Упорядковуємо колонки, повертаємо індекс, застосовуємо наявний поріг і додаємо лінію.

ordered_columns = []

for hour in hours_list:

    ordered_columns.extend(
        [
            (hour, "Single"),
            (hour, "Multi"),
            (hour,"Total"),
            (hour, "Rate")
        ]
    )

ordered_columns.extend(
    [
        ("Total", "Single"),
        ("Total", "Multi"),
        ("Total", "Quantity"),
        ("Total", "Valid hours"),
        ("Total", "Ignored hours"),
        ("Total", "Worked hours"),
        ("Total", "Break slot"),
        ("Total", "Break gap min"),
        ("Total", "Rate")
    ]
)

station_hourly = (
    station_hourly
    .reindex(
        columns=ordered_columns
    )
)

station_report = (
    station_hourly
    .reset_index()
)

min_quantity = 100

station_report_display = station_report[
    station_report[("Total", "Quantity")] > min_quantity
]

station_report_display = (
    station_report_display
    .sort_values(
        by=[
            ("Packing station", ""),
            ("Operator", "")
        ]
    )
)

station_to_line= (data[  
    ["Packing station",
     "Line"]
    ].drop_duplicates()
    .set_index("Packing station")["Line"]
)

station_report_display[("Line","")]= (
    station_report_display[("Packing station", "")]
    .map(station_to_line)
)



# ======================================================================
# 14. СТАТИСТИКА ПО СТАНЦІЯХ І ПРАЦІВНИКАХ
# ======================================================================
# Рахуємо обсяги, частки та кількість операторів за наявними правилами.

pack_type_quantity = (
    data.groupby(
        "Pack Type"
    )["Quantity of scanned"]
    .sum()
)

stats_total_single = pack_type_quantity["Single"]

stats_total_multi = pack_type_quantity["Multi"]

stats_total_quantity = pack_type_quantity.sum()

print(f"Total Single packs: {stats_total_single}")

print(f"Total Multi packs: {stats_total_multi}")

procent_single = (
    stats_total_single
    / stats_total_quantity
    * 100
)

procent_multi = (
    stats_total_multi
    / stats_total_quantity
    * 100
)

pm_single = (
    data.loc[
        (data["Pack Type"] == "Single")
        & (data["Line"] == "Pack Machine"),
        "Quantity of scanned"
    ]
    .sum()
)

manual_single = (
    stats_total_single
    - pm_single
)

print(f"Manual total single packed: {manual_single}")

print(f"Pack Machine packed: {pm_single}")

print(f"Percentage of Single packs: {procent_single:.2f}%")

print(f"Percentage of Multi packs: {procent_multi:.2f}%")


# До показників працівників включаємо лише тих,
# хто зробив понад 100 юнітів за зміну
employee_pack_type = employee_hourly.loc[
    employee_hourly[("Total", "Quantity")] > 100,
    [
        ("Total", "Single"),
        ("Total", "Multi")
    ]
].copy()

employee_pack_type.columns = ["Single", "Multi"]

employee_pack_type["Type of pack"]="Multi"

employee_pack_type["Total"] = employee_pack_type["Single"] + employee_pack_type["Multi"]

for element in employee_pack_type.index:
    if employee_pack_type.loc[element, "Single"]/120 > employee_pack_type.loc[element, "Multi"]/260:
        employee_pack_type.loc[element, "Type of pack"]="Single"

employee_work_level= (
    data
    .groupby("Employee ID")["Work level"]
    .first()
)
employee_pack_type["Work level"] = (

    employee_pack_type.index.map(employee_work_level)
)
# Один рядок у employee_pack_type — один працівник
count_veteran = (
    employee_pack_type["Work level"] == "Veteran"
).sum()

count_week1 = (
    employee_pack_type["Work level"] == "Week 1"
).sum()

count_week2 = (
    employee_pack_type["Work level"] == "Week 2"
).sum()

count_week3 = (
    employee_pack_type["Work level"] == "Week 3"
).sum()

print(f"Veteran: {count_veteran} os")
print(f"Week 1: {count_week1} os")
print(f"Week 2: {count_week2} os")
print(f"Week 3: {count_week3} os")
#=======================BREAKTIME========================


avg_breaktime=(
    employee_hourly.loc[ 
    employee_hourly[("Total", "Break gap min")]>0,
    ("Total","Break gap min")
].mean()
)

print("AVG Breaktime", avg_breaktime)



#===================================================
#    СТАТИСТИКА WEEK/ Для Dashboard
#===================================================
# ТЕСТ: призначення напрямку за найбільшим внеском у норму
# Беремо всіх працівників наявної employee_pack_type.
group_operations = data.loc[
    data["Employee ID"].isin(employee_pack_type.index)
    & data["Hour"].isin(hours_list)
].copy()

group_operations["Work type"] = "Multi"
single_mask = group_operations["Pack Type"] == "Single"
pm_mask = group_operations["Line"] == "Pack Machine"

group_operations.loc[single_mask & ~pm_mask, "Work type"] = "Single"
group_operations.loc[single_mask & pm_mask, "Work type"] = "PM"

# Зберігаємо норми чинного розрахунку: Multi 260,
# Single на PLA 315, решта Single 120 (включно з PL00).
group_operations["Base norm"] = 260.0
group_operations.loc[single_mask, "Base norm"] = 120.0
group_operations.loc[
    single_mask & pm_mask,
    "Base norm"
] = 315.0

group_operations["Contribution"] = (
    group_operations["Quantity of scanned"]
    / group_operations["Base norm"]
)

# Коефіцієнт Week однаковий для трьох напрямків однієї людини,
# тому його ділення не змінює вибір найбільшого внеску.
work_contributions = (
    group_operations
    .groupby(["Employee ID", "Work type"])["Contribution"]
    .sum()
    .unstack("Work type", fill_value=0)
    .reindex(index=employee_pack_type.index,
             columns=["Single", "PM", "Multi"], fill_value=0)
    .fillna(0)
)

# При точній рівності пріоритет: Single, PM, Multi.
assigned_work_type = work_contributions.idxmax(axis=1)
assigned_work_type = assigned_work_type.where(
    work_contributions.max(axis=1) > 0
)

employee_pack_type["Work type"] = assigned_work_type

employee_hourly[("Work type", "")] = (
    employee_pack_type["Work type"]
)


single_operator_ids = (
    employee_pack_type[
        employee_pack_type["Type of pack"] == "Single"
    ]
    .index
)
#========================================================
# Працівники, яких включаємо в таблицю dashboard
dashboard_employees = employee_pack_type.loc[
    employee_pack_type["Work level"].isin(
        ["Veteran", "Week 1", "Week 2", "Week 3"]
    )
].copy()

# Зіставляємо UPH за Employee ID
dashboard_employees["UPH"] = (
    dashboard_employees.index.map(
        employee_hourly[("Total", "UPH")]
    )
)
# ==========================================================
# RATE ПРАЦІВНИКІВ З УРАХУВАННЯМ ФАКТИЧНОГО ТИПУ РОБОТИ
# ==========================================================

level_factors = {
    "Day1": 0.55,
    "Week 1": 0.65,
    "Week 2": 0.70,
    "Week 3": 0.85,
    "Veteran": 1.00,
    "Unknown": 1.00
}

# Обсяг фактично виконаної роботи з урахуванням норм
rate_operations = data.loc[
    data["Employee ID"].isin(employee_hourly.index)
    & data["Hour"].isin(hours_list)
].copy()



# Базова норма Multi
rate_operations["Base norm"] = 260.0

# Базова норма ручного Single
rate_operations.loc[
    rate_operations["Pack Type"] == "Single",
    "Base norm"
] = 120.0

# Single на PLA має окрему норму
rate_operations.loc[
    (rate_operations["Pack Type"] == "Single")
    & (rate_operations["Line"] == "Pack Machine"),
    "Base norm"
] = 315.0

rate_operations["Norm"] = (
    rate_operations["Base norm"]
    * rate_operations["Work level"].map(level_factors)
)

rate_operations["Norm hours"] = (
    rate_operations["Quantity of scanned"]
    / rate_operations["Norm"]
)

# Збираємо результати всіх станцій за працівником і годиною
rate_hourly = (
    rate_operations
    .groupby(["Employee ID", "Hour"])[
        ["Quantity of scanned", "Norm hours"]
    ]
    .sum()
)
# Розгортаємо нормо-години один раз
employee_norm_hourly = (
    rate_hourly["Norm hours"]
    .unstack("Hour")
    .reindex(
        index=employee_hourly.index,
        columns=hours_list
    )
)

for hour in hours_list:

    valid_mask = (
        (employee_hourly[(hour, "Total")] > 15)
        & (employee_hourly[(hour, "Worked hours")] > 0)
    )

    employee_hourly[(hour, "Rate")] = float("nan")

    employee_hourly.loc[
        valid_mask,
        (hour, "Rate")
    ] = (
        employee_norm_hourly.loc[valid_mask, hour]
        / employee_hourly.loc[
            valid_mask,
            (hour, "Worked hours")
        ]
        * 100
    )

# Як у cal_rate: години з кількістю <= 15
# не додають виконану норму до Rate
valid_hours_mask = (
    rate_hourly["Quantity of scanned"] > 15
)

employee_norm_hours = (
    rate_hourly.loc[valid_hours_mask, "Norm hours"]
    .groupby(level="Employee ID")
    .sum()
)
for hour in hours_list:

    # Нормо-години працівників за поточну годину
    hour_norm_hours = (
        rate_hourly["Norm hours"]
        .unstack("Hour")
        .reindex(
            index=employee_hourly.index,
            columns=hours_list
        )[hour]
    )

    hour_worked_time = employee_hourly[(hour, "Worked hours")]

    # Rate визначаємо лише для годин із понад 15 юнітами
    valid_rate_mask = (
        (employee_hourly[(hour, "Total")] > 15)
        & (hour_worked_time > 0)
    )

    employee_hourly[(hour, "Rate")] = float("nan")

    employee_hourly.loc[
        valid_rate_mask,
        (hour, "Rate")
    ] = (
        hour_norm_hours.loc[valid_rate_mask]
        / hour_worked_time.loc[valid_rate_mask]
        * 100
    )
# Виконана норма кожного працівника за зміну
employee_hourly[("Total", "Norm hours")] = (
    employee_hourly.index.to_series()
    .map(employee_norm_hours)
    .fillna(0)
)

# Загальний Mixed Rate
employee_hourly[("Total", "Rate")] = (
    employee_hourly[("Total", "Norm hours")]
    / employee_hourly[("Total", "Worked hours")].replace(
        0, float("nan")
    )
    * 100
)

dashboard_employees["Norm hours"] = (
    dashboard_employees.index.to_series()
    .map(employee_norm_hours)
    .fillna(0)
)

dashboard_employees["Worked hours"] = (
    dashboard_employees.index.map(
        employee_hourly[("Total", "Worked hours")]
    )
)

# Нульовий час замінюємо пропуском, щоб не ділити на нуль
dashboard_employees["Rate"] = (
    dashboard_employees["Norm hours"]
    / dashboard_employees["Worked hours"].replace(0, float("nan"))
    * 100
)

# ==========================================================
# ТАБЛИЦІ DASHBOARD: HC, UPH, RATE
# ==========================================================

work_levels = ["Veteran", "Week 1", "Week 2", "Week 3"]
work_types = ["Multi", "Single", "PM"]
metrics = ["HC", "AVG UPH", "AVG Rate"]

dashboard_source = dashboard_employees.loc[
    dashboard_employees["Work level"].isin(work_levels)
    & dashboard_employees["Work type"].isin(work_types)
].copy()

# ==========================================================
# ПОГОДИННИЙ HC: MULTI / MANUAL SINGLE / PM
# ==========================================================


hc_operator_per_hour=(data.loc[
    data["Employee ID"].isin(employee_pack_type.index)&
    data["Hour"].isin(hours_list)
].copy()
)
hc_operator_per_hour["Hourly Work type"]="Multi"

hc_pm_mask = (
    (hc_operator_per_hour["Pack Type"] == "Single")
    & (hc_operator_per_hour["Line"] == "Pack Machine")
)

hc_manual_single_mask = (
    (hc_operator_per_hour["Pack Type"] == "Single")
    & (hc_operator_per_hour["Line"] != "Pack Machine")
)

hc_operator_per_hour.loc[
    hc_pm_mask,
    "Hourly Work type"
]="PM"

hc_operator_per_hour.loc[
    hc_manual_single_mask,
    "Hourly Work type"
]="Manual Single"



hc_per_hour_groupby=(
    hc_operator_per_hour
    .groupby(["Employee ID","Hour","Hourly Work type"])
    ["Quantity of scanned"]
    .sum()
    .reset_index()
)

multiple_work_types =hc_per_hour_groupby.groupby([
    "Employee ID",
    "Hour"
]).size()

multiple_work_types= multiple_work_types[multiple_work_types>1]


hc_per_hour_groupby["Base norm"]=260
hc_per_hour_groupby.loc[
        hc_per_hour_groupby["Hourly Work type"]=="PM",
        "Base norm"
]=315
hc_per_hour_groupby.loc[
        hc_per_hour_groupby["Hourly Work type"]=="Manual Single",
        "Base norm"
]=120




hc_per_hour_groupby["Contribution"]= hc_per_hour_groupby["Quantity of scanned"]/hc_per_hour_groupby["Base norm"]

hc_per_hour_groupby2=(
    hc_per_hour_groupby
    .groupby(["Employee ID","Hour"])
    ["Quantity of scanned"]
    .sum()
)
hc_per_hour_groupby2=hc_per_hour_groupby2[hc_per_hour_groupby2>15]

hc_per_hour_groupby=hc_per_hour_groupby.set_index(["Employee ID","Hour"])

hc_valid_hours = hc_per_hour_groupby.index.isin(hc_per_hour_groupby2.index)

hc_best_type_index = (hc_per_hour_groupby.loc[
    hc_valid_hours
])

hc_best_type_index = hc_best_type_index.reset_index()

hc_best_type=(
    hc_best_type_index
    .groupby(["Employee ID","Hour"])
    ["Contribution"]
    .idxmax()
)
    

hc_winner=hc_best_type_index.loc[hc_best_type]

hc_final=(
    hc_winner
    .groupby(["Hour","Hourly Work type"])
    .size()
    .unstack("Hour")

)

hc_final.loc["Total"]=hc_final.sum(axis=0)
hc_final=hc_final.reindex(columns=hours_list)
print(hc_final)


# 1. Кількість працівників
week_operator_count = (
    dashboard_source
    .groupby(["Work level", "Work type"])
    .size()
    .unstack("Work type", fill_value=0)
    .reindex(
        index=work_levels,
        columns=work_types,
        fill_value=0
    )
)

# Unknown враховуємо тільки в HC Veteran → Multi
unknown_hc = (
    employee_pack_type["Work level"] == "Unknown"
).sum()

week_operator_count.loc["Veteran","Multi"]+= unknown_hc

week_operator_count["Total"] = week_operator_count.sum(axis=1)
week_operator_count.loc["Total"] = week_operator_count.sum(axis=0)


# 2. Середній особистий UPH
week_avg_uph = (
    dashboard_source
    .groupby(["Work level", "Work type"])["UPH"]
    .mean()
    .unstack("Work type")
    .reindex(index=work_levels, columns=work_types)
)

# Підсумки рахуємо з працівників, а не із середніх клітинок
week_avg_uph["Total"] = (
    dashboard_source.groupby("Work level")["UPH"].mean()
)

week_avg_uph.loc["Total"] = (
    dashboard_source.groupby("Work type")["UPH"].mean()
)

week_avg_uph.loc["Total", "Total"] = dashboard_source["UPH"].mean()


# 3. Середній особистий Rate
week_avg_rate = (
    dashboard_source
    .groupby(["Work level", "Work type"])["Rate"]
    .mean()
    .unstack("Work type")
    .reindex(index=work_levels, columns=work_types)
)

week_avg_rate["Total"] = (
    dashboard_source.groupby("Work level")["Rate"].mean()
)

week_avg_rate.loc["Total"] = (
    dashboard_source.groupby("Work type")["Rate"].mean()
)

week_avg_rate.loc["Total", "Total"] = dashboard_source["Rate"].mean()


# 4. Об'єднання готових таблиць
dashboard_report = (
    pd.concat(
        {
            "HC": week_operator_count,
            "AVG UPH": week_avg_uph,
            "AVG Rate": week_avg_rate
        },
        names=["Metric"]
    )
    .swaplevel("Metric", "Work level")
    .reindex(
        index=pd.MultiIndex.from_product(
            [work_levels + ["Total"], metrics],
            names=["Work level", "Metric"]
        ),
        columns=work_types + ["Total"]
    )
    .round(2)
)


#==========================
#        Rate Zone Dashboard
#==========================
rate95 = (dashboard_source["Rate" ]>=95).sum()
rate80= ((dashboard_source["Rate"]<95)&(dashboard_source["Rate"]>=80)).sum()
rate60= ((dashboard_source["Rate"]<80)&(dashboard_source["Rate"]>=60)).sum()
rateless60 = (dashboard_source["Rate" ]<60).sum()

#count_total
share_ratezone=dashboard_source["Rate"].count()

share_rate95= round(rate95/share_ratezone*100,2) 
share_rate80= round(rate80/share_ratezone*100,2)
share_rate60= round(rate60/share_ratezone*100,2)
share_less60= round(rateless60/share_ratezone*100,2)

check_ratezone=share_rate95+share_rate80+share_rate60+share_less60

rate_zone_report= pd.DataFrame({
    "Rate zone":[ ">=95%","95-80%","80-60%","<60%"],
    "Operators":[rate95,rate80,rate60,rateless60],
    "shape":[share_rate95,share_rate80,share_rate60,share_less60]
    } )

print(rate_zone_report)
print("Сума частинок",check_ratezone )


#====================================================================
# ПІДРАХУНОК ПРАЦІВНИКІВ ПО ТИПУ ПАКУВАННЯ
#====================================================================
# Кількість операторів за остаточно призначеними групами
count_pm_operators = (
    employee_pack_type["Work type"] == "PM"
).sum()

count_manual_single_operator = (
    employee_pack_type["Work type"] == "Single"
).sum()

count_single_operators = (
    count_pm_operators + count_manual_single_operator
)

percent_pm_manual= (pm_single/(pm_single+manual_single)*100)

print(f"Percent pm value: {percent_pm_manual:.2f}%")

print(f"Manual single opertors: {count_manual_single_operator}")

count_multi_operators = (
   employee_pack_type
   [employee_pack_type["Type of pack"] == "Multi"]
    .shape[0]
)

count_total_operators = employee_pack_type.shape[0]

print(f"Total number of operators: {count_total_operators}")

print(f"Number of operators with more Multi packs: {count_multi_operators}")
print(f"Manual single opertors: {count_manual_single_operator}")

print(f"Pack Machine operators: {count_pm_operators}")



# ======================================================================
# 15. ПІДГОТОВКА ПОКАЗНИКІВ ДЛЯ ЗВІТУ ПО ЛІНІЯХ
# ======================================================================
# Готуємо середні Rate ліній, погодинні Manual Single та показники працівників ручного пакування.

rate_columns=[]

for hour in hours_list:
    rate_columns.append((hour,"Rate"))

test_line_rate= ( 
    station_report_display
    .groupby(("Line",""))
    [rate_columns]
    .mean()
    .round(2)
    
)

manual_single_hourly = (
    data.loc[
        (data["Pack Type"] == "Single")
        & (data["Line"] != "Pack Machine")
    ]
    .groupby("Hour")["Quantity of scanned"]
    .sum()

)


manual_single_workers = station_report_display.loc[
    station_report_display[("Employee ID", "")].isin(single_operator_ids)
    & (station_report_display[("Line", "")] != "Pack Machine")
]

manual_single_rate = (
    manual_single_workers[rate_columns]
    .mean()
    .round(2)
)

manual_single_total_rate = (
    manual_single_workers[("Total", "Rate")]
    .mean()
    .round(2)
)



# ======================================================================
# 16. ГРУПУВАННЯ ТА ЗВЕДЕНА ТАБЛИЦЯ ЛІНІЙ
# ======================================================================
# Створюємо звіт по лініях і рядок Manual single. Послідовність присвоєнь усередині блоку збережена.

line_grouped = (
    data
    .groupby(
        [
            "Line",
            "Hour",
            "Pack Type"
        ],
        dropna=False
    )["Quantity of scanned"]
    .sum()
    .reset_index()
)

line_report = (
    line_grouped
    .pivot_table(
        index="Line",
        columns=[
            "Hour",
            "Pack Type"
        ],
        values="Quantity of scanned",
        aggfunc="sum",
        fill_value=0
    )
)

for hour in hours_list:

    if (
        hour,
        "Single"
    ) not in line_report.columns:

        line_report[
            (hour, "Single")
        ] = 0


    if (
        hour,
        "Multi"
    ) not in line_report.columns:

        line_report[
            (hour, "Multi")
        ] = 0

line_report.loc["Manual single"]=0

for hour in hours_list:
    line_report.loc[
        "Manual single",
        (hour, "Single")
    ] = manual_single_hourly.get(hour, 0)

line_report.loc[ 
    "Manual single",
    ("Total","Single")
    ]=manual_single_hourly.sum()



# ======================================================================
# 17. ПІДСУМКИ ТА RATE ПО ЛІНІЯХ
# ======================================================================
# Заповнюємо підсумкові кількості та середні Rate без зміни формул.

line_report[
    ("Total", "Single")
] = 0.0

line_report[
    ("Total", "Multi")
] = 0.0

line_report[
    ("Total", "Quantity")
] = 0.0

line_report[ 
    ("Total","Rate")
]=0

for index, row in line_report.iterrows():

    total_single = 0
    total_multi = 0


    for hour in hours_list:

        single_quantity = row.get(
            (hour, "Single"),
            0
        )

        multi_quantity = row.get(
            (hour, "Multi"),
            0
        )


        if pd.isna(single_quantity):
            single_quantity = 0

        if pd.isna(multi_quantity):
            multi_quantity = 0


        total_single += (
            float(single_quantity)
        )

        total_multi += (
            float(multi_quantity)
        )


    line_report.loc[
        index,
        ("Total", "Single")
    ] = total_single


    line_report.loc[
        index,
        ("Total", "Multi")
    ] = total_multi


    line_report.loc[
        index,
        ("Total", "Quantity")
    ] = (
        total_single
        + total_multi
    )

line_report[rate_columns] = test_line_rate[rate_columns]

total_line_rate= (
    station_report_display
    .groupby(("Line", ""))[
    [("Total","Rate")]]
    .mean()
    .round(2)
    
)

line_report[("Total", "Rate")] = total_line_rate[("Total", "Rate")]

print(
    "AFTER MERGE:",
    data.loc[
        data["Pack Type"] == "Single",
        "Quantity of scanned"
    ].sum()
)



# ======================================================================
# 18. ОФОРМЛЕННЯ ЗВІТУ ПО ЛІНІЯХ
# ======================================================================
# Встановлюємо порядок колонок та записуємо Rate для Manual single.

line_columns = []
# Загальна кількість юнітів за кожну годину
for hour in hours_list:
    line_report[(hour, "Total")] = (
        line_report[(hour, "Single")]
        + line_report[(hour, "Multi")]
    )

for hour in hours_list:

    line_columns.extend(
        [
            (hour, "Single"),
            (hour, "Multi"),
            (hour,"Total"),
            (hour,'Rate')
        ]
    )

line_columns.extend(
    [
        ("Total", "Single"),
        ("Total", "Multi"),
        ("Total", "Quantity"),
        ("Total", "Rate")
    ]
)

line_report = (
    line_report
    .reindex(
        columns=line_columns
    )
)

for hour in hours_list:

    line_report.loc[
        "Manual single",
        (hour, "Rate")
    ] = manual_single_rate.get(
        (hour, "Rate"),
        pd.NA
    )

    line_report.loc[
    "Manual single",
    ("Total", "Rate")
] = manual_single_total_rate

# ======================================================================
# HC ПО ЛІНІЯХ — ОСТАННЯ ДОСТУПНА ГОДИНА
# ======================================================================

last_available_hour = hours_list[-1]

# Беремо активних працівників останньої години
# прямо з уже готового погодинного HC
last_hour_employee_ids = (
    hc_winner.loc[
        hc_winner["Hour"] == last_available_hour,
        "Employee ID"
    ]
)

# Беремо з data всі операції цих працівників за останню годину
last_hour_report = (
    data.loc[
        (data["Hour"] == last_available_hour)
        & (data["Employee ID"].isin(last_hour_employee_ids))
    ]
    .copy()
)

# Рахуємо, скільки кожен працівник зробив на кожній Line
last_hour_report_groupby = (
    last_hour_report
    .groupby(["Employee ID", "Line", "Work level"])
    ["Quantity of scanned"]
    .sum()
    .reset_index()
)

# Якщо працівник був на кількох Line —
# знаходимо Line з найбільшою Quantity
last_hour_winner_station = (
    last_hour_report_groupby
    .groupby("Employee ID")
    ["Quantity of scanned"]
    .idxmax()
)



# Один працівник = одна основна Line
hour_winner = (
    last_hour_report_groupby
    .loc[last_hour_winner_station]
    .copy()
)
hc_work_level_mask = (
    hour_winner["Work level"].isin(work_levels)
)

line_hc_report=(
    hour_winner.loc[hc_work_level_mask]
    .groupby(["Work level","Line"])
    .size()
    .unstack("Line",fill_value=0)
    .reindex(index=work_levels, fill_value=0)
)
line_hc_report.loc["Total"]=line_hc_report.sum(axis=0)
print("Test " ,line_hc_report)

libiao_map = {
    "PB300002": 1,
    "PB300003": 2,
    "PB300004": 3,
    "PB300005": 4,
    "PB300007": 6,
    "PB300008": 7,
    "PB300009": 8,
    "PB300011": 9,
    "PB300012": 10,
    "PB300013": 11,
    "PB300014": 12
}

Libiao1=[]
for line in line_hc_report.columns:
    get1=libiao_map.get(line,"")
    Libiao1.append(get1)

Libiao = pd.DataFrame(
    [Libiao1],
    columns=line_hc_report.columns,
    index=["Libiao"]
)

# ======================================================================
# 19. ПІДСУМКОВА ТАБЛИЦЯ СТАТИСТИКИ
# ======================================================================
# Збираємо вже розраховані показники в таблицю Metric / Value.

statistics_report = pd.DataFrame({
    "Metric": [
        "Total Quantity",
        "Total Single",
        "Total Multi",
        "Single %",
        "Multi %",
        "Pack Machine Single",
        "Manual Single",
        "PM %",
        "Total Operators",
        "Single Operators",
        "Pack Machine Operators",
        "Manual Single Operators",
        "Multi Operators",
        "AVG Breaktime"
    ],

    "Value": [
        stats_total_quantity,
        stats_total_single,
        stats_total_multi,
        round(procent_single, 2),
        round(procent_multi, 2),
        pm_single,
        manual_single,
        round(percent_pm_manual, 2),
        count_total_operators,
        count_single_operators,
        count_pm_operators,
        count_manual_single_operator,
        count_multi_operators,
        avg_breaktime
    ]
})

# Описові колонки на початку звіту
employee_columns_order = [
    ("Operator", ""),
    ("Work level", ""),
    ("Work type", ""),
    ("Last station", "")
]

# Погодинні показники в порядку зміни
for hour in hours_list:
    employee_columns_order.extend([
        (hour, "Single"),
        (hour, "Multi"),
        (hour, "Total"),
        (hour, "Worked hours"),
        (hour, "Rate")
    ])

# Підсумки за зміну
employee_columns_order.extend([
    ("Total", "Single"),
    ("Total", "Multi"),
    ("Total", "Quantity"),
    ("Total", "Worked hours"),
    ("Total", "UPH"),
    ("Total", "Rate"),
    ("Total", "Break slot"),
    ("Total", "Break gap min")
])

employee_report = employee_hourly.loc[
    employee_pack_type.index,
    employee_columns_order
].copy()

# ======================================================================
# 20. ЕКСПОРТ В EXCEL
# ======================================================================

output_file = "Pack_report.xlsx"

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    # =========================
    # BY OPERATOR
    # =========================
    write_table(
        writer,
        employee_report,
        "ByOperator",
        startrow=0,
        startcol=0,
        index=True
    )

    # =========================
    # BY LINE — основний звіт
    # =========================
    write_table(
        writer,
        line_report,
        "By Line",
        startrow=0,
        startcol=0,
        index=True
    )

    # =========================
    # BY LINE — LIBIAO
    # =========================
    write_table(
        writer,
        Libiao,
        "By Line",
        startrow=22,
        startcol=3,
        index=True,
        header=False,
        format_headers=False
    )

    # =========================
    # BY LINE — CURRENT HC
    # =========================
    write_table(
        writer,
        line_hc_report,
        "By Line",
        startrow=23,
        startcol=3,
        index=True
    )

    # =========================
    # STATISTICS
    # =========================
    write_table(
        writer,
        statistics_report,
        "Statistics",
        startrow=0,
        startcol=0,
        index=False
    )

    # =========================
    # DASHBOARD — основна таблиця
    # =========================
    write_table(
        writer,
        dashboard_report,
        "Dashboard",
        startrow=0,
        startcol=0,
        index=True
    )

    # =========================
    # DASHBOARD — RATE ZONE
    # =========================
    write_table(
        writer,
        rate_zone_report,
        "Dashboard",
        startrow=0,
        startcol=7,
        index=True
    )

    # =========================
    # DASHBOARD — HOURLY HC
    # =========================
    write_table(
        writer,
        hc_final,
        "Dashboard",
        startrow=8,
        startcol=7,
        index=True
    )


print(f"Звіт створено: {output_file}")