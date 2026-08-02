# vybe-test: python/py_builtins_adv/test_py_builtins_any_all_short_circuit
# origin: languages/python/tests/python/test_py_builtins_adv.rs

nums = [1, 3, 5, 7]
print(all(n % 2 != 0 for n in nums))
print(any(n > 5 for n in nums))
print(all(n > 5 for n in nums))
print(any(n % 2 == 0 for n in nums))
print(any([]))
print(all([]))  # vacuously true
