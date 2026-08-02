# vybe-test: python/walrus_extended/walrus_zip_unpack
# origin: languages/python/tests/python/test_walrus_extended.rs

print([ (a := x) + (b := y) for x, y in zip([1, 2], [10, 20]) ])
