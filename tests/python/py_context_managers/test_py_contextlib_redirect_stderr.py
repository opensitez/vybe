# vybe-test: python/py_context_managers/test_py_contextlib_redirect_stderr
# origin: languages/python/tests/python/test_py_context_managers.rs

import io, sys
from contextlib import redirect_stderr

err_buf = io.StringIO()
with redirect_stderr(err_buf):
    sys.stderr.write("Custom error message\n")

print("Captured stderr:", err_buf.getvalue().strip())
