# vybe-test: python/dict_methods_spec/dict_setdefault_list_compile
# origin: languages/python/tests/python/test_dict_methods_spec.rs
# vybe-test-mode: compile

d = {}
d.setdefault('items', []).append(1)
