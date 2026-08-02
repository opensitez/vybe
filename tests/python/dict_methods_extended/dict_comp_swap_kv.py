# vybe-test: python/dict_methods_extended/dict_comp_swap_kv
# origin: languages/python/tests/python/test_dict_methods_extended.rs

print({v: k for k, v in [('a', 1), ('b', 2)]})
