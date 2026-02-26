from unittest.mock import MagicMock

from convert_docx_to_pdf import convert_one_word


def test_success(tmp_path):
    docx = tmp_path / "report.docx"
    docx.write_bytes(b"fake")
    word_app = MagicMock()
    pdf_path, error = convert_one_word(docx, word_app)

    assert error is None
    assert pdf_path == docx.with_suffix(".pdf")
    word_app.Documents.Open.assert_called_once_with(str(docx))
    doc_mock = word_app.Documents.Open.return_value
    doc_mock.SaveAs.assert_called_once_with(str(pdf_path), FileFormat=17)
    doc_mock.Close.assert_called_once()


def test_open_exception(tmp_path):
    docx = tmp_path / "bad.docx"
    docx.write_bytes(b"fake")
    word_app = MagicMock()
    word_app.Documents.Open.side_effect = Exception("COM error 0x800")

    pdf_path, error = convert_one_word(docx, word_app)
    assert pdf_path is None
    assert "COM error 0x800" in error


def test_saveas_exception(tmp_path):
    docx = tmp_path / "fail.docx"
    docx.write_bytes(b"fake")
    word_app = MagicMock()
    word_app.Documents.Open.return_value.SaveAs.side_effect = PermissionError("locked")

    pdf_path, error = convert_one_word(docx, word_app)
    assert pdf_path is None
    assert "locked" in error
