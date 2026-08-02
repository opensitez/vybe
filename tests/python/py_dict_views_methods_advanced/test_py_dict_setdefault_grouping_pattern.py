# vybe-test: python/py_dict_views_methods_advanced/test_py_dict_setdefault_grouping_pattern
# origin: languages/python/tests/python/test_py_dict_views_methods_advanced.rs

words = ["apple", "ant", "banana", "bear", "cat"]
grouped = {}
for w in words:
    grouped.setdefault(w[0], []).append(w)

print(sorted(grouped["a"]))
print(sorted(grouped["b"]))
