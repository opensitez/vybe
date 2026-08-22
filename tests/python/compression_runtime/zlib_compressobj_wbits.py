# vybe-test: python/compression_runtime/zlib_compressobj_wbits
# origin: languages/python/tests/python/test_compression_runtime.rs

import zlib
zlib.compressobj(wbits=-zlib.MAX_WBITS)
