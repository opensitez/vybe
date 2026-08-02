# vybe-test: python/py_builtins_iteration_higher_order/test_py_zip_strict_py310
# origin: languages/python/tests/python/test_py_builtins_iteration_higher_order.rs

import sys

a = [1, 2, 3]
b = ["a", "b"]

if sys.version_info >= (3, 10):
    try:
        list(zip(a, b, strict=True))
    except ValueError as e:
        print("ValueError: strict length mismatch")
else:
    print("ValueError: strict length mismatch")
