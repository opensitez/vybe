# vybe-test: python/pythonic_idioms/list_comp_double_filter
# origin: languages/python/tests/python/test_pythonic_idioms.rs

[x for x in range(10) if x % 2 == 0 if x % 3 == 0]
