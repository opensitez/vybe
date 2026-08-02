# vybe-test: python/python_keyword_module/test_keyword_kwlist_all_are_keywords
# origin: languages/python/tests/python/test_python_keyword_module.rs

import keyword
print(all(keyword.iskeyword(k) for k in keyword.kwlist))
