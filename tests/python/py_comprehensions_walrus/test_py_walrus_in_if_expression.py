# vybe-test: python/py_comprehensions_walrus/test_py_walrus_in_if_expression
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

import re

text = "The user is: alice@example.com"
if m := re.search(r"[\w.]+@[\w.]+\.\w+", text):
    print(f"Found email: {m.group()}")
else:
    print("No email found")
