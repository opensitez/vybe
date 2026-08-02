# vybe-test: python/bytes_methods_extended/bytes_decode_error
# origin: languages/python/tests/python/test_bytes_methods_extended.rs
# vybe-test-mode: compile

try:
    b'\xff'.decode('ascii')
except UnicodeDecodeError:
    pass
