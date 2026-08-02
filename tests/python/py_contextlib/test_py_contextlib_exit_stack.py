# vybe-test: python/py_contextlib/test_py_contextlib_exit_stack
# origin: languages/python/tests/python/test_py_contextlib.rs

from contextlib import ExitStack, contextmanager

log = []

@contextmanager
def res(name):
    log.append(f"open:{name}")
    yield name
    log.append(f"close:{name}")

with ExitStack() as stack:
    for name in ["A", "B", "C"]:
        stack.enter_context(res(name))
    log.append("using")

print(log)
