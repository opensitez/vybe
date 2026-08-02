# vybe-test: python/py_contextlib_resource_management/test_py_contextlib_redirect_stderr_capture
# origin: languages/python/tests/python/test_py_contextlib_resource_management.rs

import io, sys
from contextlib import redirect_stderr

buf = io.StringIO()
with redirect_stderr(buf):
    print("ERROR MSG", file=sys.stderr)

print(buf.getvalue().strip())
