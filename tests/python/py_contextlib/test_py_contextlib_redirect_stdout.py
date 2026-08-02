# vybe-test: python/py_contextlib/test_py_contextlib_redirect_stdout
# origin: languages/python/tests/python/test_py_contextlib.rs

from contextlib import redirect_stdout
import io

buf = io.StringIO()
with redirect_stdout(buf):
    print("captured output")
    print(42)

print(buf.getvalue().strip())
print("back to real stdout")
