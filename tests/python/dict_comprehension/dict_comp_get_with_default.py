# vybe-test: python/dict_comprehension/dict_comp_get_with_default
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {k: k for k in range(2)}
print(d.get(9, 'missing'))
