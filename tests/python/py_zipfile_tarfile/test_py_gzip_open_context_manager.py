# vybe-test: python/py_zipfile_tarfile/test_py_gzip_open_context_manager
# origin: languages/python/tests/python/test_py_zipfile_tarfile.rs

import gzip, tempfile, os

with tempfile.NamedTemporaryFile(suffix=".gz", delete=False) as f:
    fname = f.name

with gzip.open(fname, "wt", encoding="utf-8") as f:
    f.write("Line 1\nLine 2\n")

with gzip.open(fname, "rt", encoding="utf-8") as f:
    lines = [l.strip() for l in f]

os.unlink(fname)
print(lines)
