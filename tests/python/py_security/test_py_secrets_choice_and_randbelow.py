# vybe-test: python/py_security/test_py_secrets_choice_and_randbelow
# origin: languages/python/tests/python/test_py_security.rs

import secrets

# secrets for cryptographic randomness
n = secrets.randbelow(100)
print(0 <= n < 100)

choice = secrets.choice("ABCDEFGHIJ")
print(choice in "ABCDEFGHIJ")

# Random password generation
alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
password = "".join(secrets.choice(alphabet) for _ in range(16))
print(len(password))
