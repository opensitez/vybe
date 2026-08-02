# vybe-test: python/filter_map_builtins/all_on_map_positive
# origin: languages/python/tests/python/test_filter_map_builtins.rs

all(x > 0 for x in map(abs, [-1, -2]))
