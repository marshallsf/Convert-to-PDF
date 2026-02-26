import pytest
from pathlib import Path


@pytest.fixture
def sample_tree(tmp_path):
    """Directory tree with .docx, .xlsx, temp (~$), unrelated, and pre-existing PDF files.

    Structure:
        tmp_path/
            doc1.docx
            sheet1.xlsx
            ~$temp.docx         (should be skipped)
            readme.txt          (unrelated)
            sub/
                doc2.docx
                doc2.pdf        (pre-existing PDF)
                sheet2.xlsx
    """
    (tmp_path / "doc1.docx").write_bytes(b"A" * 10)
    (tmp_path / "sheet1.xlsx").write_bytes(b"B" * 10)
    (tmp_path / "~$temp.docx").write_bytes(b"C" * 10)
    (tmp_path / "readme.txt").write_text("hello")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "doc2.docx").write_bytes(b"D" * 10)
    (sub / "doc2.pdf").write_bytes(b"E" * 10)
    (sub / "sheet2.xlsx").write_bytes(b"F" * 10)
    return tmp_path


@pytest.fixture
def duplicate_tree(tmp_path):
    """Directory tree with known duplicates for testing find_duplicates.

    Structure:
        tmp_path/
            a.txt       content=b"same"
            b.txt       content=b"same"     (duplicate of a.txt)
            c.txt       content=b"different"
            sub/
                d.txt   content=b"same"     (duplicate of a.txt)
                e.txt   content=b"diff"     (same size as "same" but different hash)
    """
    (tmp_path / "a.txt").write_bytes(b"same")
    (tmp_path / "b.txt").write_bytes(b"same")
    (tmp_path / "c.txt").write_bytes(b"different")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "d.txt").write_bytes(b"same")
    (sub / "e.txt").write_bytes(b"diff")
    return tmp_path


@pytest.fixture
def empty_dir(tmp_path):
    """An empty directory for negative tests."""
    return tmp_path
