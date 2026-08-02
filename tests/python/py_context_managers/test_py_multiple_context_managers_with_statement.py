# vybe-test: python/py_context_managers/test_py_multiple_context_managers_with_statement
# origin: languages/python/tests/python/test_py_context_managers.rs

class Dummy:
    def __init__(self, tag):
        self.tag = tag
    def __enter__(self):
        print(f"enter {self.tag}")
        return self.tag
    def __exit__(self, *args):
        print(f"exit {self.tag}")

with Dummy("A") as a, Dummy("B") as b:
    print(f"inside {a} {b}")
