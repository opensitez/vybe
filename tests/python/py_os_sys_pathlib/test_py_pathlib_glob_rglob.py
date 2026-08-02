# vybe-test: python/py_os_sys_pathlib/test_py_pathlib_glob_rglob
# origin: languages/python/tests/python/test_py_os_sys_pathlib.rs

import tempfile, os
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    base = Path(tmpdir)
    (base / "a.py").write_text("")
    (base / "b.py").write_text("")
    (base / "c.txt").write_text("")

    py_files = sorted([p.name for p in base.glob("*.py")])
    all_files = sorted([p.name for p in base.rglob("*") if p.is_file()])
    print(py_files)
    print(all_files)
