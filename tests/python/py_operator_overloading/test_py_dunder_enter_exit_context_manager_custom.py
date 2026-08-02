# vybe-test: python/py_operator_overloading/test_py_dunder_enter_exit_context_manager_custom
# origin: languages/python/tests/python/test_py_operator_overloading.rs

class Transaction:
    def __enter__(self):
        print("BEGIN")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            print("ROLLBACK")
            return True  # suppress
        print("COMMIT")
        return False

with Transaction():
    print("WORK")

with Transaction():
    print("WORK FAIL")
    raise ValueError("error")
