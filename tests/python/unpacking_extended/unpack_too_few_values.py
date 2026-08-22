# vybe-test: python/unpacking_extended/unpack_too_few_values
# origin: languages/python/tests/python/test_unpacking_extended.rs

try:
 a, b, c = [1, 2]
except ValueError:
 pass
