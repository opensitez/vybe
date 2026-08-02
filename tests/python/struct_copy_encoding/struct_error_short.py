# vybe-test: python/struct_copy_encoding/struct_error_short
# origin: languages/python/tests/python/test_struct_copy_encoding.rs

import struct
try:
 struct.unpack('i', b'\x01')
 print('ok')
except struct.error:
 print('short')
