# vybe-test: python/memoryview_extended/memoryview_from_bytes_readonly_assign
# origin: languages/python/tests/python/test_memoryview_extended.rs

mv = memoryview(b'abc')
try:
 mv[0] = 1
 print('ok')
except TypeError:
 print('err')
