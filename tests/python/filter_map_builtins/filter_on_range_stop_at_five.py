# vybe-test: python/filter_map_builtins/filter_on_range_stop_at_five
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(lambda x: x < 3, range(5)))
