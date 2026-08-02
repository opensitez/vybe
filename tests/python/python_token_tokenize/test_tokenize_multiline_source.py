# vybe-test: python/python_token_tokenize/test_tokenize_multiline_source
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = "a = 1\nb = 2\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
names = [t.string for t in toks if t.type == token.NAME]
print(sorted(names))
