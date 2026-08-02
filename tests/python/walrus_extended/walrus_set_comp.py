# vybe-test: python/walrus_extended/walrus_set_comp
# origin: languages/python/tests/python/test_walrus_extended.rs

print({(v := x % 2) for x in range(4)})
