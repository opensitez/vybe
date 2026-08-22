# vybe-test: python/dict_methods_spec/dict_comprehension_transform_compile
# origin: languages/python/tests/python/test_dict_methods_spec.rs

d = {k.upper(): v * 10 for k, v in [('a', 1), ('b', 2)]}
