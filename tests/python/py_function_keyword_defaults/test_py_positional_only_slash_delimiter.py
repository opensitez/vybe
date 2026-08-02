# vybe-test: python/py_function_keyword_defaults/test_py_positional_only_slash_delimiter
# origin: languages/python/tests/python/test_py_function_keyword_defaults.rs

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

def format_point(x, y, /, label="Point"):
    return f"{label}({x}, {y})"

__check(__line(format_point(10, 20)), "Point(10, 20)")
__check(__line(format_point(10, 20, label="P1")), "P1(10, 20)")
try:
    format_point(x=10, y=20)
except TypeError:
    __check(__line("TypeError: positional-only"), "TypeError: positional-only")
