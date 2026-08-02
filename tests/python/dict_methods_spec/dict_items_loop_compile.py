# vybe-test: python/dict_methods_spec/dict_items_loop_compile
# origin: languages/python/tests/python/test_dict_methods_spec.rs
# vybe-test-mode: compile

d = {'a': 1, 'b': 2}
for key, value in d.items():
    print(key, value)
