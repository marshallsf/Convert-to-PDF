import pytest
from unittest.mock import patch, MagicMock

from convert_docx_to_pdf import main_cli


# --- Validation and early exits ---

def test_invalid_directory(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: "/nonexistent/path")
    with patch("builtins.print") as mock_print, \
         pytest.raises(SystemExit) as exc_info:
        main_cli()
    assert exc_info.value.code == 1
    assert any("not a valid directory" in str(c) for c in mock_print.call_args_list)


def test_no_files_found(tmp_path, monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: str(tmp_path))
    with patch("builtins.print"), pytest.raises(SystemExit) as exc_info:
        main_cli()
    assert exc_info.value.code == 0


def test_strips_quotes(tmp_path, monkeypatch):
    (tmp_path / "f.docx").write_bytes(b"x")
    quoted = f'"{tmp_path}"'
    inputs = iter([quoted, "n"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    with patch("builtins.print"), \
         patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch("convert_docx_to_pdf.shutil.which", return_value="/usr/bin/soffice"), \
         patch("convert_docx_to_pdf.convert_one_libreoffice",
               return_value=(tmp_path / "f.pdf", None)):
        main_cli()  # Should not raise — quotes were stripped


# --- Linux / LibreOffice path ---

def test_linux_soffice_missing(tmp_path, monkeypatch):
    (tmp_path / "f.docx").write_bytes(b"x")
    monkeypatch.setattr("builtins.input", lambda _: str(tmp_path))
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch("convert_docx_to_pdf.shutil.which", return_value=None), \
         patch("builtins.print"), \
         pytest.raises(SystemExit) as exc_info:
        main_cli()
    assert exc_info.value.code == 1


def test_linux_converts_docx_and_xlsx(tmp_path, monkeypatch):
    (tmp_path / "a.docx").write_bytes(b"x")
    (tmp_path / "b.xlsx").write_bytes(b"x")
    inputs = iter([str(tmp_path), "n"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    converted = []

    def fake_lo(path):
        converted.append(path.name)
        return (path.with_suffix(".pdf"), None)

    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch("convert_docx_to_pdf.shutil.which", return_value="/usr/bin/soffice"), \
         patch("convert_docx_to_pdf.convert_one_libreoffice", side_effect=fake_lo), \
         patch("builtins.print"):
        main_cli()
    assert "a.docx" in converted
    assert "b.xlsx" in converted


def test_linux_skips_existing_pdf(tmp_path, monkeypatch):
    (tmp_path / "a.docx").write_bytes(b"x")
    (tmp_path / "a.pdf").write_bytes(b"pdf")
    inputs = iter([str(tmp_path), "n"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch("convert_docx_to_pdf.shutil.which", return_value="/usr/bin/soffice"), \
         patch("convert_docx_to_pdf.convert_one_libreoffice") as mock_lo, \
         patch("builtins.print") as mock_print:
        main_cli()
    mock_lo.assert_not_called()
    assert any("Skipped" in str(c) for c in mock_print.call_args_list)


def test_linux_conversion_failure(tmp_path, monkeypatch):
    (tmp_path / "a.docx").write_bytes(b"x")
    inputs = iter([str(tmp_path), "n"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch("convert_docx_to_pdf.shutil.which", return_value="/usr/bin/soffice"), \
         patch("convert_docx_to_pdf.convert_one_libreoffice",
               return_value=(None, "crash")), \
         patch("builtins.print") as mock_print:
        main_cli()
    assert any("FAILED" in str(c) for c in mock_print.call_args_list)
    assert any("Failed: 1" in str(c) for c in mock_print.call_args_list)


# --- Windows / COM path ---

def test_windows_comtypes_import_error(tmp_path, monkeypatch):
    (tmp_path / "f.docx").write_bytes(b"x")
    monkeypatch.setattr("builtins.input", lambda _: str(tmp_path))
    import builtins
    orig = builtins.__import__

    def mock_import(name, *a, **kw):
        if name == "comtypes":
            raise ImportError("No module named 'comtypes'")
        return orig(name, *a, **kw)

    with patch("convert_docx_to_pdf.sys.platform", "win32"), \
         patch("builtins.__import__", side_effect=mock_import), \
         patch("builtins.print"), \
         pytest.raises(SystemExit) as exc_info:
        main_cli()
    assert exc_info.value.code == 1


def test_windows_com_creation_fails(tmp_path, monkeypatch):
    (tmp_path / "f.docx").write_bytes(b"x")
    monkeypatch.setattr("builtins.input", lambda _: str(tmp_path))
    mock_comtypes = MagicMock()
    mock_comtypes.client.CreateObject.side_effect = Exception("Word not installed")
    with patch("convert_docx_to_pdf.sys.platform", "win32"), \
         patch.dict("sys.modules", {
             "comtypes": mock_comtypes,
             "comtypes.client": mock_comtypes.client,
         }), \
         patch("builtins.print"), \
         pytest.raises(SystemExit) as exc_info:
        main_cli()
    assert exc_info.value.code == 1
    mock_comtypes.CoUninitialize.assert_called()


def test_windows_dispatches_word_and_excel(tmp_path, monkeypatch):
    (tmp_path / "a.docx").write_bytes(b"x")
    (tmp_path / "b.xlsx").write_bytes(b"x")
    inputs = iter([str(tmp_path), "n"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    mock_comtypes = MagicMock()
    with patch("convert_docx_to_pdf.sys.platform", "win32"), \
         patch.dict("sys.modules", {
             "comtypes": mock_comtypes,
             "comtypes.client": mock_comtypes.client,
         }), \
         patch("convert_docx_to_pdf.convert_one_word",
               return_value=(tmp_path / "a.pdf", None)) as mock_word, \
         patch("convert_docx_to_pdf.convert_one_excel",
               return_value=(tmp_path / "b.pdf", None)) as mock_excel, \
         patch("builtins.print"):
        main_cli()
    mock_word.assert_called_once()
    mock_excel.assert_called_once()


# --- Duplicate prompt ---

def test_dedup_prompt_yes(tmp_path, monkeypatch):
    (tmp_path / "a.docx").write_bytes(b"x")
    inputs = iter([str(tmp_path), "y"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch("convert_docx_to_pdf.shutil.which", return_value="/usr/bin/soffice"), \
         patch("convert_docx_to_pdf.convert_one_libreoffice",
               return_value=(tmp_path / "a.pdf", None)), \
         patch("convert_docx_to_pdf._cli_remove_duplicates") as mock_dedup, \
         patch("builtins.print"):
        main_cli()
    mock_dedup.assert_called_once()
