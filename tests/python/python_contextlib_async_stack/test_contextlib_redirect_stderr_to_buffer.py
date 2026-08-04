# vybe-test: python/python_contextlib_async_stack/test_contextlib_redirect_stderr_to_buffer
# origin: languages/python/tests/python/test_python_contextlib_async_stack.rs

import contextlib, io, sys

buf = io.StringIO()
with contextlib.redirect_stderr(buf):
    sys.stderr.write("error message\n")

print(buf.getvalue().strip())
