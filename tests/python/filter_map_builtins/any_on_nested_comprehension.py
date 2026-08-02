# vybe-test: python/filter_map_builtins/any_on_nested_comprehension
# origin: languages/python/tests/python/test_filter_map_builtins.rs

any(x > 3 for row in [[1, 2], [4, 5]] for x in row)
