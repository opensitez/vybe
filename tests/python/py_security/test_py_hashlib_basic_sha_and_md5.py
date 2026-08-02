# vybe-test: python/py_security/test_py_hashlib_basic_sha_and_md5
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

import hashlib

h = hashlib.sha256(b"hello world")
__check(__line(h.hexdigest()), "b94d27b9934d3e08a52e52d7da7dabfac484efe04294e576f3c521f5dc8cdf2")
__check(__line(h.digest_size), "32")  # bytes
__check(__line(len(h.hexdigest())), "64")  # hex chars = 2 * bytes

md5 = hashlib.md5(b"test")
__check(__line(len(md5.hexdigest())), "32")
