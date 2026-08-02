# vybe-test: python/list_methods_extended/list_comp_nested_flatten_manual
# origin: languages/python/tests/python/test_list_methods_extended.rs

print([y for row in [[1, 2], [3]] for y in row])
