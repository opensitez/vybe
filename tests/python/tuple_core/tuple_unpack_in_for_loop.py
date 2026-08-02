# vybe-test: python/tuple_core/tuple_unpack_in_for_loop
# origin: languages/python/tests/python/test_tuple_core.rs

total = 0
for a, b in [(1, 2), (3, 4)]:
 total += a + b
print(total)
