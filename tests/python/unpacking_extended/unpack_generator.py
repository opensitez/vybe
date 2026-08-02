# vybe-test: python/unpacking_extended/unpack_generator
# origin: languages/python/tests/python/test_unpacking_extended.rs

a, b = (x for x in [1, 2])
print(a, b)
