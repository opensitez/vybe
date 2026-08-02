# vybe-test: python/dict_comprehension/dict_comp_keys_sorted
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {x: x for x in [3, 1, 2]}
print(sorted(d.keys()))
