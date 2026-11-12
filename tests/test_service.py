from io import BytesIO
from zipfile import ZipFile

import pytest

from benchmark.service import validate_upload


def minimal_package():
    target = BytesIO()
    with ZipFile(target, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/workbook.xml", "<workbook/>")
    return target.getvalue()


def test_rejects_renamed_non_excel_upload():
    with pytest.raises(ValueError, match="无法读取"):
        validate_upload(b"not a spreadsheet", "broken.xlsx")


def test_macro_content_is_rejected_even_with_xlsx_extension():
    buffer = BytesIO()
    with ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/workbook.xml", "<workbook/>")
        archive.writestr("xl/vbaProject.bin", b"not executable")
    with pytest.raises(ValueError, match="不含宏"):
        validate_upload(buffer.getvalue(), "renamed.xlsx")


def test_accepts_workbook_package_and_rejects_wrong_extension():
    validate_upload(minimal_package(), "book.xlsx")
    with pytest.raises(ValueError, match="xlsx"):
        validate_upload(minimal_package(), "book.xls")
