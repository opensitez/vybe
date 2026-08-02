# vybe-test: python/filter_map_builtins/any_on_filter_nonzero
# origin: languages/python/tests/python/test_filter_map_builtins.rs

any(x > 0 for x in [-1, 0, 2])
