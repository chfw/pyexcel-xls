from io import BytesIO
from datetime import date

import xlwt
import pytest
from pyexcel_xls import get_data


@pytest.mark.parametrize("source", ["stream", "file"])
@pytest.mark.parametrize("book_mode", [0, 1])
@pytest.mark.parametrize("override", ["default", None, 0, 1])
def test_date_mode_override(tmp_path, source, book_mode, override):
    workbook = xlwt.Workbook()
    workbook.dates_1904 = bool(book_mode)
    sheet = workbook.add_sheet("dates")
    sheet.write(0, 0, 61, xlwt.easyxf(num_format_str="YYYY-MM-DD"))
    if source == "stream":
        target = BytesIO()
        workbook.save(target)
        target.seek(0)
    else:
        target = str(tmp_path / "dates.xls")
        workbook.save(target)
    options = {} if override == "default" else {"date_mode": override}

    data = get_data(target, library="pyexcel-xls", **options)

    mode = book_mode if override in ("default", None) else override
    expected = date(1900, 3, 1) if mode == 0 else date(1904, 3, 2)
    assert data["dates"] == [[expected]]


def test_date_mode_override_handles_early_serial_on_every_sheet():
    workbook = xlwt.Workbook()
    for name in ("first", "second"):
        sheet = workbook.add_sheet(name)
        sheet.write(0, 0, 1, xlwt.easyxf(num_format_str="YYYY-MM-DD"))
    target = BytesIO()
    workbook.save(target)
    target.seek(0)

    data = get_data(target, date_mode=1, library="pyexcel-xls")

    assert data["first"] == [[date(1904, 1, 2)]]
    assert data["second"] == [[date(1904, 1, 2)]]
