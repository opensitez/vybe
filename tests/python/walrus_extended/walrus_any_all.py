# vybe-test: python/walrus_extended/walrus_any_all
# origin: languages/python/tests/python/test_walrus_extended.rs

print(any((v := x) > 2 for x in [1, 3]))
