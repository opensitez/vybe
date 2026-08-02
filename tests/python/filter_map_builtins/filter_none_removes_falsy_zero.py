# vybe-test: python/filter_map_builtins/filter_none_removes_falsy_zero
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(None, [0, 1, 2, '', 3]))
