# vybe-test: python/stdlib_compile_extended/lzma_compress
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import lzma
lzma.compress(b'hi')
