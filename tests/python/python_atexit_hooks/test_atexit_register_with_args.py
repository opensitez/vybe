# vybe-test: python/python_atexit_hooks/test_atexit_register_with_args
# origin: languages/python/tests/python/test_python_atexit_hooks.rs

import atexit

log = []

def record(msg, count):
    for _ in range(count):
        log.append(msg)

atexit.register(record, "bye", 3)
atexit._run_exitfuncs()
print(log)
