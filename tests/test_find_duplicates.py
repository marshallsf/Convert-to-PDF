from pathlib import Path

import convert_docx_to_pdf as mod
from convert_docx_to_pdf import find_duplicates


def test_identifies_duplicates(duplicate_tree):
    results = find_duplicates(duplicate_tree)
    assert len(results) == 1
    kept, dupes = results[0]
    assert len(dupes) == 2
    # All three should be the files with content b"same"
    all_paths = [kept] + dupes
    all_names = {p.name for p in all_paths}
    assert all_names == {"a.txt", "b.txt", "d.txt"}


def test_unique_files_not_reported(duplicate_tree):
    results = find_duplicates(duplicate_tree)
    all_paths = []
    for kept, dupes in results:
        all_paths.append(kept)
        all_paths.extend(dupes)
    all_names = {p.name for p in all_paths}
    assert "c.txt" not in all_names


def test_same_size_different_hash(duplicate_tree):
    results = find_duplicates(duplicate_tree)
    all_paths = []
    for kept, dupes in results:
        all_paths.append(kept)
        all_paths.extend(dupes)
    all_names = {p.name for p in all_paths}
    # e.txt has 4 bytes like "same" but content "diff" — should not be a duplicate
    assert "e.txt" not in all_names


def test_empty_dir(empty_dir):
    assert find_duplicates(empty_dir) == []


def test_no_duplicates(tmp_path):
    (tmp_path / "x.txt").write_bytes(b"one")
    (tmp_path / "y.txt").write_bytes(b"two")
    assert find_duplicates(tmp_path) == []


def test_skips_temp_files(tmp_path):
    (tmp_path / "a.txt").write_bytes(b"data")
    (tmp_path / "~$a.txt").write_bytes(b"data")
    results = find_duplicates(tmp_path)
    assert results == []  # only one non-temp file, no dupe group possible


def test_progress_callback(duplicate_tree):
    calls = []
    find_duplicates(duplicate_tree, progress_cb=lambda s, t: calls.append((s, t)))
    assert len(calls) > 0
    # Final call should have scanned == total
    assert calls[-1][0] == calls[-1][1]


def test_progress_callback_none(duplicate_tree):
    # Should not raise when progress_cb is None (the default)
    find_duplicates(duplicate_tree, progress_cb=None)


def test_oserror_on_stat(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_bytes(b"data")
    (tmp_path / "b.txt").write_bytes(b"data")
    original_stat = Path.stat
    # Track calls to b.txt stat — first call is from is_file() (let it pass),
    # second call is from the size-grouping f.stat().st_size (make it fail).
    b_call_count = {"n": 0}

    def bad_stat(self, *a, **kw):
        if self.name == "b.txt":
            b_call_count["n"] += 1
            if b_call_count["n"] > 1:
                raise OSError("permission denied")
        return original_stat(self, *a, **kw)

    monkeypatch.setattr(Path, "stat", bad_stat)
    results = find_duplicates(tmp_path)
    # b.txt skipped at stat phase, only a.txt in its size group -> no dupe
    assert results == []


def test_oserror_on_hash(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_bytes(b"data")
    (tmp_path / "b.txt").write_bytes(b"data")
    original_hash = mod._file_hash

    def bad_hash(path, *a, **kw):
        if path.name == "b.txt":
            raise OSError("read error")
        return original_hash(path, *a, **kw)

    monkeypatch.setattr(mod, "_file_hash", bad_hash)
    results = find_duplicates(tmp_path)
    # b.txt hash skipped, only a.txt hashed -> no dupe group
    assert results == []


def test_keeps_first_sorted(tmp_path):
    (tmp_path / "zzz.txt").write_bytes(b"dup")
    (tmp_path / "aaa.txt").write_bytes(b"dup")
    results = find_duplicates(tmp_path)
    assert len(results) == 1
    kept, dupes = results[0]
    assert kept.name == "aaa.txt"
    assert dupes[0].name == "zzz.txt"
