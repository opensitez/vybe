# vybe-test: python/python_token_tokenize/test_tokenize_detect_encoding_utf8
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io
src = b"# -*- coding: utf-8 -*-\nx = 1\n"
toks = list(tokenize.tokenize(io.BytesIO(src).readline))
encodings = [t.string for t in toks if t.type == tokenize.ENCODING]
print(encodings)
