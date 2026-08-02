# vybe-test: python/python_colorsys_conversions/test_colorsys_hsv_roundtrip
# origin: languages/python/tests/python/test_python_colorsys_conversions.rs

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

import colorsys
r0, g0, b0 = 0.1, 0.4, 0.7
h, s, v = colorsys.rgb_to_hsv(r0, g0, b0)
r1, g1, b1 = colorsys.hsv_to_rgb(h, s, v)
__check(__line(round(r1, 5) == round(r0, 5)), "True")
__check(__line(round(g1, 5) == round(g0, 5)), "True")
