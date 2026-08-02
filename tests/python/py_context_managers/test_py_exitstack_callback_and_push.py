# vybe-test: python/py_context_managers/test_py_exitstack_callback_and_push
# origin: languages/python/tests/python/test_py_context_managers.rs

from contextlib import ExitStack

def cleanup(item):
    print(f"Cleanup {item}")

with ExitStack() as stack:
    stack.callback(cleanup, "Item1")
    stack.callback(cleanup, "Item2")
    print("Done stack tasks")
