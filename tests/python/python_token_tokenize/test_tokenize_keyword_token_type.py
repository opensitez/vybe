# vybe-test: python/python_token_tokenize/test_tokenize_keyword_token_type
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = "if True:\n    pass\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
names = [t.string for t in toks if t.type == token.NAME]
print("if" in names)
print("True" in names)
print("pass" in names)
