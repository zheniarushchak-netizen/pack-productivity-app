"""Завантаження файлів та зрозуміла перевірка їхнього формату."""
from io import BytesIO
import pandas as pd

PACK_COLUMNS = ["Package No.", "Item quantity classification name", "Operator",
                "Packing station", "Operation time", "Quantity of scanned"]
EMPLOYEE_COLUMNS = ["Employee ID", "Create time"]

def read_employees(employee_bytes):
    try:
        employees = pd.read_excel(
            BytesIO(employee_bytes),
            engine="openpyxl"
        )
        
    except Exception as exc:
        raise ValueError(
            "Не вдалося прочитати довідник працівників."
        ) from exc

    missing = [
        column
        for column in EMPLOYEE_COLUMNS
        if column not in employees.columns
    ]

    if missing:
        raise ValueError(
            f"У довіднику відсутні колонки: {', '.join(missing)}"
        )

    if employees.empty:
        raise ValueError(
            "Довідник працівників не містить рядків."
        )

    return employees[EMPLOYEE_COLUMNS].copy()


def read_inputs(pack_bytes, employee_bytes):
    frames = []

    for content, columns, label in [
        (pack_bytes, PACK_COLUMNS, "Packing operations"),
        (employee_bytes, EMPLOYEE_COLUMNS, "Employee directory")
    ]:
        try:
            with pd.ExcelFile(
                BytesIO(content),
                engine="openpyxl"
            ) as workbook:

                matching_sheets = []

                # Check column names without loading all rows
                for sheet_name in workbook.sheet_names:
                    headers = pd.read_excel(
                        workbook,
                        sheet_name=sheet_name,
                        nrows=0
                    )

                    if all(
                        column in headers.columns
                        for column in columns
                    ):
                        matching_sheets.append(sheet_name)

                if not matching_sheets:
                    raise ValueError(
                        f"{label}: no worksheet contains "
                        "all required columns."
                    )

                if len(matching_sheets) > 1:
                    raise ValueError(
                        f"{label}: multiple matching worksheets: "
                        f"{', '.join(matching_sheets)}. "
                        "Keep one source worksheet."
                    )

                frame = pd.read_excel(
                    workbook,
                    sheet_name=matching_sheets[0]
                )

        except ValueError:
            raise

        except Exception as exc:
            raise ValueError(
                f"{label}: failed to read the Excel file."
            ) from exc
        
        missing = [c for c in columns if c not in frame.columns]
        if missing:
            raise ValueError(f"{label}: відсутні колонки: {', '.join(missing)}")
        if frame.empty:
            raise ValueError(f"{label}: файл не містить рядків даних.")
        frames.append(frame[columns].copy())
    return tuple(frames)


def input_notes(pack, employees):
    notes = []
    ids = employees["Employee ID"]
    if ids.duplicated().any():
        notes.append("У довіднику є повторні Employee ID. Оригінальне об’єднання може дублювати операції; рядки не видалено.")
    if pack["Operator"].astype(str).str.extract(r"\((\d+)\)")[0].isna().any():
        notes.append("Частина операторів не має ID у дужках. Правила їх обробки залишені як в оригіналі.")
    dates = pd.to_datetime(pack["Operation time"], errors="coerce")
    if dates.isna().any():
        notes.append("Некоректний час операцій буде відкинуто оригінальним алгоритмом.")
    if dates.dropna().dt.normalize().nunique() > 2:
        notes.append("Файл охоплює більше двох дат. Рекомендовано завантажувати одну зміну: оригінал групує за годиною, а не за датою.")
    return notes
