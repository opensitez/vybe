# vybe-test: python/py_os_pathlib_filesystem_traversal/test_py_pathlib_glob_rglob_pattern_matching
# origin: languages/python/tests/python/test_py_os_pathlib_filesystem_traversal.rs

import tempfile
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    root = Path(tmpdir)
    (root / "a.py").write_text("code_a")
    (root / "b.txt").write_text("text_b")
    sub = root / "sub"
    sub.mkdir()
    (sub / "c.py").write_text("code_c")

    py_files = sorted([p.name for p in root.rglob("*.py")])
    print(py_files)
