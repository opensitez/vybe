# vybe-test: python/bytes_buffer_spec/memoryview_tobytes_compile
# origin: languages/python/tests/python/test_bytes_buffer_spec.rs

m = memoryview(b'abc')
b = m.tobytes()
