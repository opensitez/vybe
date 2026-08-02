# vybe-test: python/py_dict_views_methods_advanced/test_py_dict_comprehension_conditional_filtering
# origin: languages/python/tests/python/test_py_dict_views_methods_advanced.rs

raw = {"a": 1, "b": 2, "c": 3, "d": 4}
evens = {k: v * 10 for k, v in raw.items() if v % 2 == 0}
print(sorted(evens.items()))
