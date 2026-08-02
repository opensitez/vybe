# vybe-test: python/python_keyword_module/test_keyword_softkwlist_is_list
# origin: languages/python/tests/python/test_python_keyword_module.rs

import keyword, sys
if hasattr(keyword, "softkwlist"):
    print(isinstance(keyword.softkwlist, list))
else:
    print(True)
