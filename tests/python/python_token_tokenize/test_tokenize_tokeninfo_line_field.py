# vybe-test: python/python_token_tokenize/test_tokenize_tokeninfo_line_field
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = "answer = 42\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
num = [t for t in toks if t.type == token.NUMBER][0]
print("42" in num.line)
