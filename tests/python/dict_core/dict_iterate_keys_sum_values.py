# vybe-test: python/dict_core/dict_iterate_keys_sum_values
# origin: languages/python/tests/python/test_dict_core.rs

d = {'a': 1, 'b': 2}
print(sum(d[k] for k in d))
