# vybe-test: python/dict_methods_spec/dict_values_loop_compile
# origin: languages/python/tests/python/test_dict_methods_spec.rs

d = {'a': 1, 'b': 2}
for value in d.values():
    print(value)
