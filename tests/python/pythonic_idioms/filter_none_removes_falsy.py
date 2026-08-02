# vybe-test: python/pythonic_idioms/filter_none_removes_falsy
# origin: languages/python/tests/python/test_pythonic_idioms.rs

list(filter(None, [0, 1, '', 'x']))
