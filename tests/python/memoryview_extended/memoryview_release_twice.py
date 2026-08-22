# vybe-test: python/memoryview_extended/memoryview_release_twice
# origin: languages/python/tests/python/test_memoryview_extended.rs

mv = memoryview(b'a')
mv.release()
