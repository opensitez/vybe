# vybe-test: python/python_keyword_module/test_keyword_kwlist_contains_common_keywords
# origin: languages/python/tests/python/test_python_keyword_module.rs

import keyword
kws = keyword.kwlist
for k in ["if", "else", "while", "return", "import", "lambda"]:
    print(k in kws)
