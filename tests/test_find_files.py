from pathlib import Path

from convert_docx_to_pdf import SUPPORTED_EXTENSIONS, find_convertible_files


def test_supported_extensions_contains_docx_and_xlsx():
    assert SUPPORTED_EXTENSIONS == {".docx", ".xlsx"}


def test_finds_docx_and_xlsx(sample_tree):
    files = find_convertible_files(sample_tree)
    names = {f.name for f in files}
    assert "doc1.docx" in names
    assert "sheet1.xlsx" in names
    assert "doc2.docx" in names
    assert "sheet2.xlsx" in names
    assert len(files) == 4


def test_skips_temp_files(sample_tree):
    files = find_convertible_files(sample_tree)
    names = {f.name for f in files}
    assert "~$temp.docx" not in names


def test_skips_unrelated_extensions(sample_tree):
    files = find_convertible_files(sample_tree)
    suffixes = {f.suffix for f in files}
    assert ".txt" not in suffixes
    assert ".pdf" not in suffixes


def test_returns_sorted(sample_tree):
    files = find_convertible_files(sample_tree)
    assert files == sorted(files)


def test_empty_dir(empty_dir):
    assert find_convertible_files(empty_dir) == []


def test_nested_deeply(tmp_path):
    deep = tmp_path / "a" / "b" / "c"
    deep.mkdir(parents=True)
    (deep / "deep.docx").write_bytes(b"x")
    files = find_convertible_files(tmp_path)
    assert len(files) == 1
    assert files[0].name == "deep.docx"
