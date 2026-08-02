# vybe-test: python/dict_methods_extended/dict_comp_nested
# origin: languages/python/tests/python/test_dict_methods_extended.rs
# vybe-test-mode: compile

d = {k: {k: k} for k in range(2)}
