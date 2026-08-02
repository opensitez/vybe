# vybe-test: python/list_methods_extended/list_comp_double_condition
# origin: languages/python/tests/python/test_list_methods_extended.rs

print([x for x in range(10) if x % 2 == 0 if x > 3])
