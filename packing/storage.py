"""Підключення до спільного сховища."""

import streamlit as st
from supabase import create_client


def check_storage():
    settings = st.secrets["supabase"]

    client = create_client(
        settings["url"],
        settings["secret_key"]
    )

    client.storage.from_(settings["bucket"]).list(
        options={"limit": 1}
    )