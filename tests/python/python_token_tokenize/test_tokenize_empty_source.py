# vybe-test: python/python_token_tokenize/test_tokenize_empty_source
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io
src = ""
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
types = [t.type for t in toks]
print(tokenize.ENDMARKER in types)
