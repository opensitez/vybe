# vybe-test: python/bytes_buffer_spec/bytearray_slice_assign_compile
# origin: languages/python/tests/python/test_bytes_buffer_spec.rs

b = bytearray(b'abc')
b[1:2] = b'XYZ'
