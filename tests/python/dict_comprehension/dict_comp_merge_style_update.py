# vybe-test: python/dict_comprehension/dict_comp_merge_style_update
# origin: languages/python/tests/python/test_dict_comprehension.rs

base = {'a': 1}
base.update({k: k for k in range(2)})
print(base)
