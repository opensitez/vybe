# vybe-test: python/py_re/test_py_re_compile_and_reuse
# origin: languages/python/tests/python/test_py_re.rs

import re

pattern = re.compile(r'\b\w{5}\b')
texts = ["Hello world", "The quick brown fox", "abc defgh"]
for t in texts:
    found = pattern.findall(t)
    print(found if found else [])
