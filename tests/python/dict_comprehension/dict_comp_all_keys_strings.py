# vybe-test: python/dict_comprehension/dict_comp_all_keys_strings
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {str(i): i for i in range(2)}
print(all(isinstance(k, str) for k in d))
