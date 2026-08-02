# vybe-test: python/compression_runtime/bz2_error
# origin: languages/python/tests/python/test_compression_runtime.rs

import bz2
try:
 bz2.decompress(b'bad')
 print('ok')
except Exception:
 print('err')
