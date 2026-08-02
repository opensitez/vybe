# vybe-test: python/python_token_tokenize/test_tokenize_error_on_invalid_token
# origin: languages/python/tests/python/test_python_token_tokenize.rs

import tokenize, io
src = "$\n"
try:
    list(tokenize.generate_tokens(io.StringIO(src).readline))
    print("no error")
except tokenize.TokenError:
    print("TokenError")
except Exception:
    print("other error")
