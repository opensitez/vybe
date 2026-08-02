# vybe-test: python/py_io/test_py_io_append_mode
# origin: languages/python/tests/python/test_py_io.rs

import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    fname = f.name
    f.write("initial\n")

with open(fname, 'a') as f:
    f.write("appended\n")

with open(fname) as f:
    lines = f.readlines()
    print([l.strip() for l in lines])

os.unlink(fname)
