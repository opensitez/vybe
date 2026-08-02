# vybe-test: python/python_base64_encodings_variants/test_base64_b32hexencode_and_decode
# origin: languages/python/tests/python/test_python_base64_encodings_variants.rs

import base64, sys
if hasattr(base64, "b32hexencode"):
    data = b"base32hex test"
    enc = base64.b32hexencode(data)
    dec = base64.b32hexdecode(enc)
    print(dec == data)
else:
    print(True)
