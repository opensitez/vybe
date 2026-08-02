# vybe-test: python/list_methods_extended/list_comp_nested_squares
# origin: languages/python/tests/python/test_list_methods_extended.rs

print([[i * j for j in range(3)] for i in range(2)])
