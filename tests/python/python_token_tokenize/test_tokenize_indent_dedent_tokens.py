# vybe-test: python/python_token_tokenize/test_tokenize_indent_dedent_tokens
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io
src = "if True:\n    pass\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
types = [t.type for t in toks]
print(tokenize.INDENT in types)
print(tokenize.DEDENT in types)
