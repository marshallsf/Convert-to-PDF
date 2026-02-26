import hashlib

import pytest

from convert_docx_to_pdf import _file_hash


def test_known_sha256(tmp_path):
    p = tmp_path / "data.bin"
    content = b"hello world"
    p.write_bytes(content)
    expected = hashlib.sha256(content).hexdigest()
    assert _file_hash(p) == expected


def test_empty_file(tmp_path):
    p = tmp_path / "empty.bin"
    p.write_bytes(b"")
    expected = hashlib.sha256(b"").hexdigest()
    assert _file_hash(p) == expected


def test_large_file_chunked(tmp_path):
    p = tmp_path / "large.bin"
    content = b"x" * 200_000  # > 65536 default buf_size
    p.write_bytes(content)
    expected = hashlib.sha256(content).hexdigest()
    assert _file_hash(p) == expected


def test_custom_buf_size(tmp_path):
    p = tmp_path / "small_buf.bin"
    content = b"abcdefgh" * 100
    p.write_bytes(content)
    expected = hashlib.sha256(content).hexdigest()
    assert _file_hash(p, buf_size=16) == expected


def test_nonexistent_raises(tmp_path):
    with pytest.raises(OSError):
        _file_hash(tmp_path / "nope.bin")
