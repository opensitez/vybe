# vybe-test: python/py_os_pathlib_filesystem_traversal/test_py_os_walk_tree_traversal
# origin: languages/python/tests/python/test_py_os_pathlib_filesystem_traversal.rs

import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "subdir")
    os.mkdir(sub)
    open(os.path.join(tmpdir, "top.txt"), "w").close()
    open(os.path.join(sub, "sub.txt"), "w").close()

    visited_files = []
    for root, dirs, files in os.walk(tmpdir):
        visited_files.extend(files)

    print(sorted(visited_files))
