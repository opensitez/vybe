# vybe-test: python/memoryview_extended/memoryview_subclass
# origin: languages/python/tests/python/test_memoryview_extended.rs
# vybe-test-mode: compile

class M(memoryview):
 pass
M(b'a')
