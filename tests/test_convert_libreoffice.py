from subprocess import TimeoutExpired
from unittest.mock import MagicMock, patch

from convert_docx_to_pdf import convert_one_libreoffice


def test_success(tmp_path):
    docx = tmp_path / "test.docx"
    docx.write_bytes(b"fake")
    pdf = tmp_path / "test.pdf"
    mock_result = MagicMock(returncode=0)

    with patch("convert_docx_to_pdf.subprocess.run", return_value=mock_result):
        pdf.write_bytes(b"fake pdf")  # simulate LibreOffice creating the file
        result_path, error = convert_one_libreoffice(docx)

    assert error is None
    assert result_path == pdf


def test_nonzero_exit_stderr(tmp_path):
    docx = tmp_path / "test.docx"
    docx.write_bytes(b"fake")
    mock_result = MagicMock(returncode=1, stderr="fatal error in soffice")

    with patch("convert_docx_to_pdf.subprocess.run", return_value=mock_result):
        result_path, error = convert_one_libreoffice(docx)

    assert result_path is None
    assert "fatal error" in error


def test_nonzero_exit_no_stderr(tmp_path):
    docx = tmp_path / "test.docx"
    docx.write_bytes(b"fake")
    mock_result = MagicMock(returncode=42, stderr="")

    with patch("convert_docx_to_pdf.subprocess.run", return_value=mock_result):
        result_path, error = convert_one_libreoffice(docx)

    assert result_path is None
    assert "42" in error


def test_pdf_not_created(tmp_path):
    docx = tmp_path / "test.docx"
    docx.write_bytes(b"fake")
    mock_result = MagicMock(returncode=0)

    with patch("convert_docx_to_pdf.subprocess.run", return_value=mock_result):
        # Don't create the PDF file — simulate LibreOffice silently failing
        result_path, error = convert_one_libreoffice(docx)

    assert result_path is None
    assert "not created" in error


def test_timeout(tmp_path):
    docx = tmp_path / "test.docx"
    docx.write_bytes(b"fake")

    with patch("convert_docx_to_pdf.subprocess.run",
               side_effect=TimeoutExpired("soffice", 120)):
        result_path, error = convert_one_libreoffice(docx)

    assert result_path is None
    assert error  # TimeoutExpired string representation


def test_generic_exception(tmp_path):
    docx = tmp_path / "test.docx"
    docx.write_bytes(b"fake")

    with patch("convert_docx_to_pdf.subprocess.run",
               side_effect=FileNotFoundError("soffice not found")):
        result_path, error = convert_one_libreoffice(docx)

    assert result_path is None
    assert "not found" in error


def test_command_args(tmp_path):
    docx = tmp_path / "test.docx"
    docx.write_bytes(b"fake")
    pdf = tmp_path / "test.pdf"
    mock_result = MagicMock(returncode=0)

    with patch("convert_docx_to_pdf.subprocess.run", return_value=mock_result) as mock_run:
        pdf.write_bytes(b"pdf")
        convert_one_libreoffice(docx)

    args, kwargs = mock_run.call_args
    assert args[0] == [
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", str(tmp_path), str(docx),
    ]
    assert kwargs["capture_output"] is True
    assert kwargs["text"] is True
    assert kwargs["timeout"] == 120
