# vybe-test: python/set_methods_extended/set_comp_nested_gen
# origin: languages/python/tests/python/test_set_methods_extended.rs

print(sorted({(x, y) for x in range(2) for y in range(2)}))
