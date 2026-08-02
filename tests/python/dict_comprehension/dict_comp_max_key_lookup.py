# vybe-test: python/dict_comprehension/dict_comp_max_key_lookup
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {i: i * i for i in range(4)}
print(d[max(d)])
