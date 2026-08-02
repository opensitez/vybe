# vybe-test: python/py_contextlib_resource_management/test_py_contextlib_contextmanager_exception_rollback
# origin: languages/python/tests/python/test_py_contextlib_resource_management.rs

from contextlib import contextmanager

@contextmanager
def db_transaction():
    print("BEGIN TRANSACTION")
    try:
        yield
        print("COMMIT")
    except ValueError:
        print("ROLLBACK")

with db_transaction():
    print("FAILING")
    raise ValueError("Query error")
