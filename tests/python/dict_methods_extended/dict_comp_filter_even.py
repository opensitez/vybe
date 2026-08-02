# vybe-test: python/dict_methods_extended/dict_comp_filter_even
# origin: languages/python/tests/python/test_dict_methods_extended.rs

print({x: x for x in range(5) if x % 2 == 0})
