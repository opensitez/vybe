# vybe-test: python/py_str_bytes_bytearray_conversions/test_py_bytearray_methods_split_replace
# origin: languages/python/tests/python/test_py_str_bytes_bytearray_conversions.rs

ba = bytearray(b"foo,bar,baz")
parts = ba.split(b",")
print([p.decode() for p in parts])

ba.replace(b"bar", b"qux")
print(ba.decode())
