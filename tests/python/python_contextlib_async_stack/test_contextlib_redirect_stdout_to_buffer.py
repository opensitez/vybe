# vybe-test: python/python_contextlib_async_stack/test_contextlib_redirect_stdout_to_buffer
# origin: languages/python/tests/python/test_python_contextlib_async_stack.rs

import contextlib, io

buf = io.StringIO()
with contextlib.redirect_stdout(buf):
    print("Line 1")
    print("Line 2")

print(buf.getvalue().strip().split("\n"))
