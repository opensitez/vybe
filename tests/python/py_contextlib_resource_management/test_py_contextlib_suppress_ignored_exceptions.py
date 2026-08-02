# vybe-test: python/py_contextlib_resource_management/test_py_contextlib_suppress_ignored_exceptions
# origin: languages/python/tests/python/test_py_contextlib_resource_management.rs

from contextlib import suppress

print("before")
with suppress(FileNotFoundError, KeyError):
    raise KeyError("missing")
    print("unreachable")
print("after")
