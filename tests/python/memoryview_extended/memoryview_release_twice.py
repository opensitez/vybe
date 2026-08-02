# vybe-test: python/memoryview_extended/memoryview_release_twice
# origin: languages/python/tests/python/test_memoryview_extended.rs
# vybe-test-mode: compile

mv = memoryview(b'a')
mv.release()
