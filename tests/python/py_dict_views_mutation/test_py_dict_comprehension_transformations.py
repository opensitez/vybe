# vybe-test: python/py_dict_views_mutation/test_py_dict_comprehension_transformations
# origin: languages/python/tests/python/test_py_dict_views_mutation.rs

scores = {"alice": 85, "bob": 92, "charlie": 78}
passed = {k.capitalize(): v for k, v in scores.items() if v >= 80}
print(sorted(passed.items()))
