# vybe-test: python/python_keyword_module/test_keyword_issoftkeyword_type_is_soft
# origin: languages/python/tests/python/test_python_keyword_module.rs

import keyword, sys
if sys.version_info >= (3, 12):
    print(keyword.issoftkeyword("type"))
else:
    print(True)
