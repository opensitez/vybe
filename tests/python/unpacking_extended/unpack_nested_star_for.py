# vybe-test: python/unpacking_extended/unpack_nested_star_for
# origin: languages/python/tests/python/test_unpacking_extended.rs

for (a, *rest), b in [((1, 2, 3), 4)]:
 print(a, len(rest), b)
