# vybe-test: python/unpacking_extended/unpack_too_many_values
# origin: languages/python/tests/python/test_unpacking_extended.rs

try:
 a, b = [1, 2, 3]
except ValueError:
 pass
