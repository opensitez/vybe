# vybe-test: python/pythonic_idioms/truthy_filter_in_comprehension
# origin: languages/python/tests/python/test_pythonic_idioms.rs

[x for x in [0, 1, 2, '', 'a'] if x]
