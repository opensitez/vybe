# vybe-test: python/dict_methods_extended/dict_comp_if_else_value
# origin: languages/python/tests/python/test_dict_methods_extended.rs

print({x: ('even' if x % 2 == 0 else 'odd') for x in range(3)})
