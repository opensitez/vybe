# vybe-test: python/walrus_extended/walrus_min_max
# origin: languages/python/tests/python/test_walrus_extended.rs

print(max((v := x) for x in [1, 5, 3]))
