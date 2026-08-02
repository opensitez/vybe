# vybe-test: python/py_time_profile/test_py_time_strftime_strptime
# origin: languages/python/tests/python/test_py_time_profile.rs

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

import time

t_struct = time.struct_time((2024, 6, 15, 12, 30, 0, 5, 167, 0))
formatted = time.strftime("%Y-%m-%d %H:%M:%S", t_struct)
__check(__line(formatted), "2024-06-15 12:30:00")

parsed = time.strptime("2024-06-15 12:30:00", "%Y-%m-%d %H:%M:%S")
__check(__line(parsed.tm_year, parsed.tm_mon, parsed.tm_mday), "2024 6 15")
