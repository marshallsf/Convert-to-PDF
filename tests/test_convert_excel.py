from unittest.mock import MagicMock

from convert_docx_to_pdf import convert_one_excel


def test_success(tmp_path):
    xlsx = tmp_path / "budget.xlsx"
    xlsx.write_bytes(b"fake")
    excel_app = MagicMock()
    pdf_path, error = convert_one_excel(xlsx, excel_app)

    assert error is None
    assert pdf_path == xlsx.with_suffix(".pdf")
    excel_app.Workbooks.Open.assert_called_once_with(str(xlsx))
    wb_mock = excel_app.Workbooks.Open.return_value
    wb_mock.ExportAsFixedFormat.assert_called_once_with(0, str(pdf_path))
    wb_mock.Close.assert_called_once_with(False)


def test_open_exception(tmp_path):
    xlsx = tmp_path / "bad.xlsx"
    xlsx.write_bytes(b"fake")
    excel_app = MagicMock()
    excel_app.Workbooks.Open.side_effect = Exception("Excel COM error")

    pdf_path, error = convert_one_excel(xlsx, excel_app)
    assert pdf_path is None
    assert "Excel COM error" in error


def test_export_exception(tmp_path):
    xlsx = tmp_path / "fail.xlsx"
    xlsx.write_bytes(b"fake")
    excel_app = MagicMock()
    excel_app.Workbooks.Open.return_value.ExportAsFixedFormat.side_effect = OSError("disk full")

    pdf_path, error = convert_one_excel(xlsx, excel_app)
    assert pdf_path is None
    assert "disk full" in error
