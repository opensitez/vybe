# vybe-test: python/python_keyword_module/test_keyword_softkwlist_all_are_soft_keywords
# origin: languages/python/tests/python/test_python_keyword_module.rs

import keyword, sys
if hasattr(keyword, "softkwlist") and hasattr(keyword, "issoftkeyword"):
    print(all(keyword.issoftkeyword(k) for k in keyword.softkwlist))
else:
    print(True)
