# vybe-test: python/python_token_tokenize/test_tokenize_basic_expression
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = "x = 1 + 2\n"
tokens = list(tokenize.generate_tokens(io.StringIO(src).readline))
names = [tok.string for tok in tokens if tok.type not in (token.ENCODING, token.ENDMARKER, token.NEWLINE, token.NL)]
print(names)
