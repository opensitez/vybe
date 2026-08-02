# vybe-test: python/py_contextlib/test_py_contextlib_redirect_stderr
# origin: languages/python/tests/python/test_py_contextlib.rs

from contextlib import redirect_stderr
import io, sys

buf = io.StringIO()
with redirect_stderr(buf):
    print("error message", file=sys.stderr)

print(buf.getvalue().strip())
