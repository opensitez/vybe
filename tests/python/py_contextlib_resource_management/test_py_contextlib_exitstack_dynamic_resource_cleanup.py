# vybe-test: python/py_contextlib_resource_management/test_py_contextlib_exitstack_dynamic_resource_cleanup
# origin: languages/python/tests/python/test_py_contextlib_resource_management.rs

from contextlib import ExitStack, contextmanager

@contextmanager
def acquire(name):
    print(f"Acquired {name}")
    try:
        yield name
    finally:
        print(f"Released {name}")

with ExitStack() as stack:
    resources = [stack.enter_context(acquire(f"R{i}")) for i in range(3)]
    print("Working with resources")
