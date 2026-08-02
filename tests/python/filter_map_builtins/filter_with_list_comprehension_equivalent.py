# vybe-test: python/filter_map_builtins/filter_with_list_comprehension_equivalent
# origin: languages/python/tests/python/test_filter_map_builtins.rs

a = list(filter(lambda x: x % 2 == 1, range(5)))
b = [x for x in range(5) if x % 2 == 1]
print(a == b)
