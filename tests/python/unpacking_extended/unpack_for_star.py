# vybe-test: python/unpacking_extended/unpack_for_star
# origin: languages/python/tests/python/test_unpacking_extended.rs

out = []
for *mid, last in [(1, 2, 3)]:
 out.append(last)
print(out)
