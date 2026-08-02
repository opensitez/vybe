# vybe-test: python/pythonic_idioms/dict_comp_with_if_filter
# origin: languages/python/tests/python/test_pythonic_idioms.rs

{k: v for k, v in [('a', 1), ('b', 2)] if v > 1}
