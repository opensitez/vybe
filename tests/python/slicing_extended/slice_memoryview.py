# vybe-test: python/slicing_extended/slice_memoryview
# origin: languages/python/tests/python/test_slicing_extended.rs

mv = memoryview(b'abcd')
bytes(mv[1:3])
