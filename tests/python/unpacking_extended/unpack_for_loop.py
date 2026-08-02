# vybe-test: python/unpacking_extended/unpack_for_loop
# origin: languages/python/tests/python/test_unpacking_extended.rs

pairs = []
for x, y in [(1, 2), (3, 4)]:
 pairs.append(x + y)
print(pairs)
