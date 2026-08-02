# vybe-test: python/py_os_sys_pathlib/test_py_os_listdir_and_scandir
# origin: languages/python/tests/python/test_py_os_sys_pathlib.rs

import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    for name in ["a.txt", "b.py", "c.md"]:
        open(os.path.join(tmpdir, name), "w").close()

    names = sorted(os.listdir(tmpdir))
    print(names)

    entries = sorted([e.name for e in os.scandir(tmpdir)])
    print(entries)
