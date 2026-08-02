# vybe-test: python/set_methods_extended/set_comp_filter
# origin: languages/python/tests/python/test_set_methods_extended.rs

print(sorted({x for x in range(6) if x % 2 == 1}))
