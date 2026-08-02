# vybe-test: python/python_token_tokenize/test_tokenize_comment_token
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io
src = "x = 1  # this is a comment\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
comments = [t.string for t in toks if t.type == tokenize.COMMENT]
print(comments)
