# vybe-test: python/stdlib_compile_extended/tokenize_tokenize
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

import tokenize
import io
tokenize.tokenize(io.BytesIO(b'1').readline)
