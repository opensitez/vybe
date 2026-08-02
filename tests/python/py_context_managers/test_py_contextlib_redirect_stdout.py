# vybe-test: python/py_context_managers/test_py_contextlib_redirect_stdout
# origin: languages/python/tests/python/test_py_context_managers.rs

import io
from contextlib import redirect_stdout

f = io.StringIO()
with redirect_stdout(f):
    print("Hello to buffer!")
    print("Second line")

print("Captured:", f.getvalue().strip())
