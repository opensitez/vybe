# vybe-test: python/py_re/test_py_re_findall_and_finditer
# origin: languages/python/tests/python/test_py_re.rs

import re

text = "The price is $10 and $25.50"
amounts = re.findall(r'\$(\d+(?:\.\d+)?)', text)
print(amounts)

positions = [(m.start(), m.group()) for m in re.finditer(r'\$\d+', text)]
print(positions)
