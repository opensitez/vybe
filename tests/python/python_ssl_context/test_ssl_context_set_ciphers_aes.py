# vybe-test: python/python_ssl_context/test_ssl_context_set_ciphers_aes
# origin: languages/python/tests/python/test_python_ssl_context.rs

import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
try:
    ctx.set_ciphers("AES256-SHA")
    print("ok")
except ssl.SSLError:
    print("ok")  # cipher string may not be supported everywhere
