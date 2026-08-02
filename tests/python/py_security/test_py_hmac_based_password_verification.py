# vybe-test: python/py_security/test_py_hmac_based_password_verification
# origin: languages/python/tests/python/test_py_security.rs

# Vybe test harness — Python.
#
# The Python counterpart of harness/go/check.go and harness/js/check.js: real
# source in the language under test, the way test262's assert.js is JavaScript.
#
# A test's verdict is its EXIT CODE. `__check` prints its diagnostic BEFORE
# raising, because an uncaught exception surfaces as `RuntimeError: [object]`
# and tells you nothing — 1,692 of testecma's 2,158 failures say exactly that.


def __line(*args):
    """One printed line, exactly as `print` composes it: str() each argument,
    joined by a single space.

    Written with a plain loop on purpose. `" ".join(str(a) for a in args)` —
    a GENERATOR EXPRESSION as a call argument — returns the empty string under
    Vybe while CPython gives "1 x", and a harness must not depend on anything
    the runtime under test gets wrong. A list comprehension works, but a loop
    needs nothing at all: no comprehension, no `join`, no `range`.
    """
    out = ""
    first = True
    for a in args:
        if not first:
            out += " "
        out += str(a)
        first = False
    return out


def __check(got, want):
    if got != want:
        print("FAIL: want [" + want + "] got [" + got + "]")
        raise Exception("assertion failed")

import hmac, hashlib, secrets

def hash_password(password: str, salt: bytes = None) -> tuple:
    if salt is None:
        salt = secrets.token_bytes(16)
    h = hmac.new(salt, password.encode(), hashlib.sha256)
    return h.hexdigest(), salt

def verify_password(password: str, stored_hash: str, salt: bytes) -> bool:
    h = hmac.new(salt, password.encode(), hashlib.sha256)
    return hmac.compare_digest(h.hexdigest(), stored_hash)

stored, salt = hash_password("my_password")
__check(__line(verify_password("my_password", stored, salt)), "True")
__check(__line(verify_password("wrong_password", stored, salt)), "False")
