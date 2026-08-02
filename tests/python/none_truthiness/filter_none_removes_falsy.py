# vybe-test: python/none_truthiness/filter_none_removes_falsy
# origin: languages/python/tests/python/test_none_truthiness.rs

list(filter(None, [0, 1, '', 'a', None]))
