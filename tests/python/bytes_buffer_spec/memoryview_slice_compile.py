# vybe-test: python/bytes_buffer_spec/memoryview_slice_compile
# origin: languages/python/tests/python/test_bytes_buffer_spec.rs
# vybe-test-mode: compile

m = memoryview(b'abcdef')
s = m[1:4]
