# vybe-test: python/bytes_core/bytes_maketrans_translate
# origin: languages/python/tests/python/test_bytes_core.rs

b'abc'.translate(bytes.maketrans(b'a', b'x'))
