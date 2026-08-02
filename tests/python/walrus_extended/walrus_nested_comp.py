# vybe-test: python/walrus_extended/walrus_nested_comp
# origin: languages/python/tests/python/test_walrus_extended.rs

print([ (a := i) + (b := 1) for i in range(2) ])
