# vybe-test: python/python_token_tokenize/test_tokenize_number_tokens
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = "42 3.14 0xFF\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
numbers = [t.string for t in toks if t.type == token.NUMBER]
print(numbers)
