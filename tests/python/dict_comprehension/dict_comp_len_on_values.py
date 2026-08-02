# vybe-test: python/dict_comprehension/dict_comp_len_on_values
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {k: len(k) for k in ['a', 'bbb']}
print(d['bbb'])
