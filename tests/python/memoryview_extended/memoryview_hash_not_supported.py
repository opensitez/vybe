# vybe-test: python/memoryview_extended/memoryview_hash_not_supported
# origin: languages/python/tests/python/test_memoryview_extended.rs

try:
 hash(memoryview(b'a'))
 print('ok')
except TypeError:
 print('err')
