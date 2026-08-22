# vybe-test: python/memoryview_extended/memoryview_subclass
# origin: languages/python/tests/python/test_memoryview_extended.rs
# `memoryview` is not an acceptable base type — asserting that rejection
# is the test.
try:
    class MV(memoryview):
        pass
except TypeError:
    pass
