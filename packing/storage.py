"""Підключення до спільного сховища."""

import streamlit as st
from supabase import create_client


def get_bucket():
    settings = st.secrets["supabase"]

    client = create_client(
        settings["url"],
        settings["secret_key"]
    )

    return client.storage.from_(settings["bucket"])


def check_storage():
    bucket = get_bucket()

    bucket.list(
        options={"limit": 1}
    )

def save_employees(employee_bytes):
    bucket = get_bucket()

    bucket.upload(
        path="employees/current.xlsx",
        file=employee_bytes,
        file_options={
            "content-type": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "upsert": "true"
        }
    )

def load_employees():
    bucket = get_bucket()

    employee_bytes = bucket.download(
        "employees/current.xlsx"
    )

    return employee_bytes