# vybe-test: python/py_os_pathlib_filesystem_traversal/test_py_os_scandir_entry_attributes
# origin: languages/python/tests/python/test_py_os_pathlib_filesystem_traversal.rs

import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    fpath = os.path.join(tmpdir, "test.txt")
    with open(fpath, "w") as f:
        f.write("content")

    with os.scandir(tmpdir) as entries:
        for entry in entries:
            print(entry.name)
            print(entry.is_file())
            print(entry.stat().st_size > 0)
