# vybe-test: python/python_reprlib_repr_limits/test_reprlib_recursive_depth_prevents_infinite_loop
# origin: languages/python/tests/python/test_python_reprlib_repr_limits.rs

import reprlib
# Deeply nested structure
lst = [None]
for _ in range(50):
    lst = [lst]
result = reprlib.repr(lst)
print(isinstance(result, str))
print("..." in result)
