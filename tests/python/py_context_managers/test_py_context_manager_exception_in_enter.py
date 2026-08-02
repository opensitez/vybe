# vybe-test: python/py_context_managers/test_py_context_manager_exception_in_enter
# origin: languages/python/tests/python/test_py_context_managers.rs

class FailingEnter:
    def __enter__(self):
        print("enter failed")
        raise RuntimeError("Enter error")
    def __exit__(self, *args):
        print("exit should NOT be called")

try:
    with FailingEnter():
        print("inside block")
except RuntimeError as e:
    print(f"Caught: {e}")
