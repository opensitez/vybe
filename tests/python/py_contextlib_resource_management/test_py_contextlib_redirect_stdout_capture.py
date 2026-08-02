# vybe-test: python/py_contextlib_resource_management/test_py_contextlib_redirect_stdout_capture
# origin: languages/python/tests/python/test_py_contextlib_resource_management.rs

import io
from contextlib import redirect_stdout

buf = io.StringIO()
with redirect_stdout(buf):
    print("Captured message 1")
    print("Captured message 2")

print(buf.getvalue().strip().splitlines())
