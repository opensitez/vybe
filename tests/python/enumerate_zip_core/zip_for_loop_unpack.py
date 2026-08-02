# vybe-test: python/enumerate_zip_core/zip_for_loop_unpack
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

s = 0
for a, b in zip([1, 2], [10, 20]):
 s += a + b
print(s)
