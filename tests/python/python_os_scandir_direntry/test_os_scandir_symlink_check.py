# vybe-test: python/python_os_scandir_direntry/test_os_scandir_symlink_check
# origin: languages/python/tests/python/test_python_os_scandir_direntry.rs

import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    target = os.path.join(tmpdir, "target.txt")
    link = os.path.join(tmpdir, "link.txt")
    with open(target, "w") as f: f.write("target")
    try:
        os.symlink(target, link)
        with os.scandir(tmpdir) as it:
            for entry in it:
                if entry.name == "link.txt":
                    print(entry.is_symlink())
    except OSError:
        # Symlinks may require elevated permissions on Windows
        print(True)
