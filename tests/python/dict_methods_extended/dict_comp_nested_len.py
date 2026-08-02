# vybe-test: python/dict_methods_extended/dict_comp_nested_len
# origin: languages/python/tests/python/test_dict_methods_extended.rs

print({k: len(k) for k in ['aa', 'bbb']})
