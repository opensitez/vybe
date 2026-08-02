# vybe-test: python/compression_runtime/zlib_error
# origin: languages/python/tests/python/test_compression_runtime.rs

import zlib
try:
 zlib.decompress(b'not zlib')
 print('ok')
except zlib.error:
 print('err')
