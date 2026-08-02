# vybe-test: python/py_context_managers/test_py_contextlib_exitstack_dynamic_contexts
# origin: languages/python/tests/python/test_py_context_managers.rs

from contextlib import ExitStack

class Tracker:
    def __init__(self, name):
        self.name = name
    def __enter__(self):
        print(f"Enter {self.name}")
        return self
    def __exit__(self, *args):
        print(f"Exit {self.name}")

with ExitStack() as stack:
    t1 = stack.enter_context(Tracker("First"))
    t2 = stack.enter_context(Tracker("Second"))
    print("Inside ExitStack")
