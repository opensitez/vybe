# vybe-test: python/python_os_scandir_direntry/test_os_walk_directory_tree
# origin: languages/python/tests/python/test_python_os_scandir_direntry.rs

import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "subdir")
    os.mkdir(sub)
    with open(os.path.join(sub, "f.txt"), "w") as f: f.write("x")

    found_dirs = []
    found_files = []
    for root, dirs, files in os.walk(tmpdir):
        found_dirs.extend(dirs)
        found_files.extend(files)

    print("subdir" in found_dirs)
    print("f.txt" in found_files)
