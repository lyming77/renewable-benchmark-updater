from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_app_requires_both_uploads_before_generating():
    script = Path(__file__).resolve().parents[1] / "streamlit_app.py"
    app = AppTest.from_file(str(script)).run(timeout=20)
    assert not app.exception
    assert app.title[0].value == "新能源交易对标更新"
    assert app.button[0].disabled
    assert app.number_input[0].value == 2026
