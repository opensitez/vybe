# vybe-test: python/unpacking_extended/unpack_nested_mismatch
# origin: languages/python/tests/python/test_unpacking_extended.rs

try:
 (a, b), c = (1, 2, 3)
except ValueError:
 pass
