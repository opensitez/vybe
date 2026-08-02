# vybe-test: python/py_builtins_iteration_higher_order/test_py_all_any_generator_expressions
# origin: languages/python/tests/python/test_py_builtins_iteration_higher_order.rs

nums = [2, 4, 6, 8, 10]
print(all(x % 2 == 0 for x in nums))
print(any(x > 5 for x in nums))
print(any(x > 100 for x in nums))
