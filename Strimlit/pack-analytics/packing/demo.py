"""Вигадані дані для тестування без робочих Excel-файлів."""
import datetime as dt
from io import BytesIO
import pandas as pd
from .inputs import PACK_COLUMNS


def demo_frames(night=False, extended=False):
    today = pd.Timestamp(dt.date.today())
    rows, employees = [], []
    for i, age in enumerate([45, 5, 12, 19, 0, 60, 8, 20, 35, 3, 14, 50]):
        employee_id = 1001 + i
        employees.append({"Employee ID": employee_id, "Create time": today - pd.Timedelta(days=age)})
        station = ["PB300002-01", "PB300003-02", "PLA-01"][i % 3]
        for offset in range(12 if extended else 11):
            base = today + pd.Timedelta(hours=(18 if night else 6) + offset)
            # A realistic gap around the default break slot.
            for minute in ([50] if offset == 4 else [5, 25, 45]):
                for kind in ["Single product", "Multiple products"]:
                    preferred = (kind == "Single product") == (i % 3 != 1)
                    quantity = (25 + i * 3) if preferred else 4
                    rows.append([f"DEMO-{len(rows)+1:05}", kind, f"Оператор {i+1:02} ({employee_id})",
                                 station, base + pd.Timedelta(minutes=minute), quantity])
        if extended:
            rows.append([f"DEMO-EXT-{i}", "Single product", f"Оператор {i+1:02} ({employee_id})",
                         station, today + pd.Timedelta(hours=28 if night else 16, minutes=55), 25])
    return pd.DataFrame(rows, columns=PACK_COLUMNS), pd.DataFrame(employees)


def as_excel(frame):
    buffer = BytesIO()
    frame.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()
