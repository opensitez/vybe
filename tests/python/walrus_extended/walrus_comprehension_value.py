# vybe-test: python/walrus_extended/walrus_comprehension_value
# origin: languages/python/tests/python/test_walrus_extended.rs

print([ (x := i + 1) for i in range(2) ])
