# vybe-test: python/filter_map_builtins/filter_map_chain_square_evens
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(map(lambda x: x * x, filter(lambda x: x % 2 == 0, range(5))))
