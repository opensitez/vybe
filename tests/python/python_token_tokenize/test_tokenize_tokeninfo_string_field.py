# vybe-test: python/python_token_tokenize/test_tokenize_tokeninfo_string_field
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io
src = '"hello world"\n'
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
strings = [t.string for t in toks if t.type == tokenize.STRING]
print(strings)
