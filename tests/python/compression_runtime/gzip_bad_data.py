# vybe-test: python/compression_runtime/gzip_bad_data
# origin: languages/python/tests/python/test_compression_runtime.rs

import gzip
try:
 gzip.decompress(b'bad')
 print('ok')
except Exception:
 print('err')
