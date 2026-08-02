# vybe-test: python/bytes_decode_encode/bytes_decode_errors_replace
# origin: languages/python/tests/python/test_bytes_decode_encode.rs

b'\xff'.decode('ascii', errors='replace')
