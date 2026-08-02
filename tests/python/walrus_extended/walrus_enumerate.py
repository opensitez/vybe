# vybe-test: python/walrus_extended/walrus_enumerate
# origin: languages/python/tests/python/test_walrus_extended.rs

print([ (i := idx) for idx, _ in enumerate(['x', 'y']) ])
