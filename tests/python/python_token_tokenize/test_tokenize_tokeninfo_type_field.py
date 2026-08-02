# vybe-test: python/python_token_tokenize/test_tokenize_tokeninfo_type_field
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io, token
src = "hello\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
names_tok = [t for t in toks if t.type == token.NAME]
print(len(names_tok) == 1)
print(names_tok[0].string)
