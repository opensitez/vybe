# vybe-test: python/py_os_sys_pathlib/test_py_os_walk_directory_tree
# origin: languages/python/tests/python/test_py_os_sys_pathlib.rs

import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "sub")
    os.makedirs(sub)
    open(os.path.join(tmpdir, "root.txt"), "w").close()
    open(os.path.join(sub, "child.txt"), "w").close()

    files_found = []
    for root, dirs, files in os.walk(tmpdir):
        for f in files:
            files_found.append(f)
    print(sorted(files_found))
