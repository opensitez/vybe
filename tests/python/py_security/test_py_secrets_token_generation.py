# vybe-test: python/py_security/test_py_secrets_token_generation
# origin: languages/python/tests/python/test_py_security.rs

import secrets

token_hex = secrets.token_hex(16)
print(len(token_hex))      # 32 hex chars for 16 bytes
print(all(c in "0123456789abcdef" for c in token_hex))

token_bytes = secrets.token_bytes(16)
print(len(token_bytes))

token_url = secrets.token_urlsafe(16)
print(len(token_url) >= 16)  # url-safe base64 is longer
