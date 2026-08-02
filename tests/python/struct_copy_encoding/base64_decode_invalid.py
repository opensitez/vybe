# vybe-test: python/struct_copy_encoding/base64_decode_invalid
# origin: languages/python/tests/python/test_struct_copy_encoding.rs

import base64
try:
 base64.b64decode(b'!!!')
 print('ok')
except Exception:
 print('bad')
