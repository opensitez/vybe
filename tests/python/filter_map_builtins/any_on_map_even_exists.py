# vybe-test: python/filter_map_builtins/any_on_map_even_exists
# origin: languages/python/tests/python/test_filter_map_builtins.rs

any(x % 2 == 0 for x in map(int, ['1', '2', '3']))
