"""Завантаження файлів та зрозуміла перевірка їхнього формату."""
from io import BytesIO
import pandas as pd

PACK_COLUMNS = ["Package No.", "Item quantity classification name", "Operator",
                "Packing station", "Operation time", "Quantity of scanned"]
EMPLOYEE_COLUMNS = ["Employee ID", "Create time"]


def read_inputs(pack_bytes, employee_bytes):
    frames = []
    for content, columns, label in [(pack_bytes, PACK_COLUMNS, "Операції пакування"),
                                    (employee_bytes, EMPLOYEE_COLUMNS, "Довідник працівників")]:
        try:
            frame = pd.read_excel(BytesIO(content), engine="openpyxl")
        except Exception as exc:
            raise ValueError(f"{label}: не вдалося прочитати Excel. Завантажте коректний .xlsx файл.") from exc
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
