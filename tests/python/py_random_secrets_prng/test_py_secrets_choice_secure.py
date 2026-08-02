# vybe-test: python/py_random_secrets_prng/test_py_secrets_choice_secure
# origin: languages/python/tests/python/test_py_random_secrets_prng.rs

import secrets

chars = "abcdefghijklmnopqrstuvwxyz0123456789"
pwd = "".join(secrets.choice(chars) for _ in range(12))
print(len(pwd))
print(all(c in chars for c in pwd))
