# vybe-test: python/python_os_scandir_direntry/test_os_walk_topdown_false
# origin: languages/python/tests/python/test_python_os_scandir_direntry.rs

import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "a", "b")
    os.makedirs(sub)
    roots = [root for root, _, _ in os.walk(tmpdir, topdown=False)]
    print(roots[0].endswith("b"))
