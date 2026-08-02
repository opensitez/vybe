# vybe-test: python/zip_patterns/zip_unpack_in_for_loop
# origin: languages/python/tests/python/test_zip_patterns.rs

s = 0
for a, b in zip([1, 2], [10, 20]):
 s += a + b
print(s)
