# vybe-test: python/py_builtins_adv/test_py_builtins_sorted_stable
# origin: languages/python/tests/python/test_py_builtins_adv.rs

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

data = [(1, "b"), (2, "a"), (1, "a"), (2, "b")]
# Stable sort: equal keys preserve original order
by_first = sorted(data, key=lambda x: x[0])
__check(__line(by_first), "[(1, 'b'), (1, 'a'), (2, 'a'), (2, 'b')]")

# Sort with multiple criteria
by_both = sorted(data, key=lambda x: (x[0], x[1]))
__check(__line(by_both), "[(1, 'a'), (1, 'b'), (2, 'a'), (2, 'b')]")
