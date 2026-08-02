# vybe-test: python/python_keyword_module/test_keyword_issoftkeyword_hard_keywords_are_not_soft
# origin: languages/python/tests/python/test_python_keyword_module.rs

import keyword, sys
if hasattr(keyword, "issoftkeyword"):
    print(keyword.issoftkeyword("def"))
    print(keyword.issoftkeyword("class"))
else:
    print(False)
    print(False)
