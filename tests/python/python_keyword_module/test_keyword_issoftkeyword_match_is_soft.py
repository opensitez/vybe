# vybe-test: python/python_keyword_module/test_keyword_issoftkeyword_match_is_soft
# origin: languages/python/tests/python/test_python_keyword_module.rs

import keyword, sys
if sys.version_info >= (3, 9):
    print(keyword.issoftkeyword("match"))
else:
    print(True)
