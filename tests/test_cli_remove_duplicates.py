from pathlib import Path
from unittest.mock import patch

from convert_docx_to_pdf import _cli_remove_duplicates


def test_no_duplicates_found(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_bytes(b"unique1")
    (tmp_path / "b.txt").write_bytes(b"unique2")
    with patch("builtins.print") as mock_print:
        _cli_remove_duplicates(tmp_path)
    assert any("No duplicate" in str(c) for c in mock_print.call_args_list)


def test_displays_report(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_bytes(b"same")
    (tmp_path / "b.txt").write_bytes(b"same")
    inputs = iter(["n"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    with patch("builtins.print") as mock_print:
        _cli_remove_duplicates(tmp_path)
    printed = " ".join(str(c) for c in mock_print.call_args_list)
    assert "Keeping" in printed
    assert "Removing" in printed


def test_confirm_yes_deletes(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_bytes(b"same")
    (tmp_path / "b.txt").write_bytes(b"same")
    monkeypatch.setattr("builtins.input", lambda _: "y")
    with patch("builtins.print"):
        _cli_remove_duplicates(tmp_path)
    # One file should have been deleted
    remaining = list(tmp_path.iterdir())
    assert len(remaining) == 1


def test_confirm_no_aborts(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_bytes(b"same")
    (tmp_path / "b.txt").write_bytes(b"same")
    monkeypatch.setattr("builtins.input", lambda _: "n")
    with patch("builtins.print") as mock_print:
        _cli_remove_duplicates(tmp_path)
    # Both files should still exist
    remaining = list(tmp_path.iterdir())
    assert len(remaining) == 2
    assert any("Aborted" in str(c) for c in mock_print.call_args_list)


def test_deletion_oserror(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_bytes(b"same")
    (tmp_path / "b.txt").write_bytes(b"same")
    monkeypatch.setattr("builtins.input", lambda _: "y")

    original_unlink = Path.unlink

    def bad_unlink(self, *a, **kw):
        if self.name == "b.txt":
            raise OSError("permission denied")
        original_unlink(self, *a, **kw)

    monkeypatch.setattr(Path, "unlink", bad_unlink)
    with patch("builtins.print") as mock_print:
        _cli_remove_duplicates(tmp_path)
    # Error message should be printed
    assert any("Could not delete" in str(c) for c in mock_print.call_args_list)
