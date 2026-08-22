# vybe-test: python/dict_methods_spec/dict_comprehension_guard_compile
# origin: languages/python/tests/python/test_dict_methods_spec.rs

d = {k: v for k, v in [('a', 1), ('b', 2)] if v > 1}
