# vybe-test: python/py_comprehensions_walrus/test_py_comprehension_with_walrus_early_exit
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

import re

emails = ["alice@example.com", "not-an-email", "bob@test.org", "broken@", "carol@domain.net"]
valid = [
    m.group()
    for email in emails
    if (m := re.fullmatch(r"[\w.]+@[\w]+\.\w+", email))
]
print(valid)
