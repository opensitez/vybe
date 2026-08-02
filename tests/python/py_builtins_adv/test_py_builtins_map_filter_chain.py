# vybe-test: python/py_builtins_adv/test_py_builtins_map_filter_chain
# origin: languages/python/tests/python/test_py_builtins_adv.rs

data = range(10)
result = list(map(lambda x: x**2, filter(lambda x: x % 2 == 0, data)))
print(result)

# Equivalent with comprehension:
result2 = [x**2 for x in data if x % 2 == 0]
print(result2)
