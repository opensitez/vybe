# vybe-test: python/python_token_tokenize/test_tokenize_fstring_token
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token, sys
src = 'f"hello {name}"\n'
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
# f-strings may be STRING or multiple tokens depending on Python version
string_toks = [t for t in toks if t.type == token.STRING or t.type == token.NAME]
print(len(string_toks) >= 1)
