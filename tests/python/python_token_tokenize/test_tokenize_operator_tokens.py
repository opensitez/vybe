# vybe-test: python/python_token_tokenize/test_tokenize_operator_tokens
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = "a + b * c\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
ops = [t.string for t in toks if t.type == token.OP]
print(ops)
