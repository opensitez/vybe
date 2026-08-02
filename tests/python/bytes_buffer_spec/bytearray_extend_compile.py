# vybe-test: python/bytes_buffer_spec/bytearray_extend_compile
# origin: languages/python/tests/python/test_bytes_buffer_spec.rs
# vybe-test-mode: compile

b = bytearray(b'a')
b.extend(b'bc')
