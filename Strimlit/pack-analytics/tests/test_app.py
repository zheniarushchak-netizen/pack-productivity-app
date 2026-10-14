from pathlib import Path
from streamlit.testing.v1 import AppTest


def test_demo_and_upload_empty_state():
    app = AppTest.from_file(str(Path(__file__).resolve().parents[1]/'app.py')).run(timeout=60)
    assert not app.exception
    assert len(app.dataframe) == 9
    assert len(app.metric) == 4
    app.toggle[0].set_value(True).run(timeout=60)
    assert not app.exception
    assert app.session_state['result']['night'] is True
    app.toggle[1].set_value(True).run(timeout=60)
    assert not app.exception
    assert app.session_state['result']['extended'] is True
    app.radio[0].set_value('Мої Excel-файли').run()
    assert not app.exception
    assert len(app.dataframe) == 0
