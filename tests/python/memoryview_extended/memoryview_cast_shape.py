# vybe-test: python/memoryview_extended/memoryview_cast_shape
# origin: languages/python/tests/python/test_memoryview_extended.rs

mv = memoryview(b'1234')
mv.cast('I')
