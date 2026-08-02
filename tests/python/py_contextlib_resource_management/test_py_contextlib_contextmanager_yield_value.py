# vybe-test: python/py_contextlib_resource_management/test_py_contextlib_contextmanager_yield_value
# origin: languages/python/tests/python/test_py_contextlib_resource_management.rs

from contextlib import contextmanager

@contextmanager
def db_transaction():
    print("BEGIN TRANSACTION")
    try:
        yield "db_connection"
        print("COMMIT")
    except Exception:
        print("ROLLBACK")
        raise

with db_transaction() as conn:
    print(f"Executing query with {conn}")
