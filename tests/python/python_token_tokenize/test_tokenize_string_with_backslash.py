# vybe-test: python/python_token_tokenize/test_tokenize_string_with_backslash
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = r'"hello\nworld"' + "\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
strings = [t.string for t in toks if t.type == token.STRING]
print(len(strings) == 1)
