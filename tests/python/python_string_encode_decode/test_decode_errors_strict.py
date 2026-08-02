# vybe-test: python/python_string_encode_decode/test_decode_errors_strict
# origin: languages/python/tests/python/test_python_string_encode_decode.rs

b = bytes([0xFF, 0xFE])
try:
    b.decode('utf-8', errors='strict')
    print("no_error")
except UnicodeDecodeError:
    print("UnicodeDecodeError")
