# vybe-test: python/py_operator_overloading/test_py_dunder_add_radd_iadd
# origin: languages/python/tests/python/test_py_operator_overloading.rs

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

class Number:
    def __init__(self, val):
        self.val = val

    def __add__(self, other):
        v = other.val if isinstance(other, Number) else other
        return Number(self.val + v)

    def __radd__(self, other):
        return self.__add__(other)

    def __iadd__(self, other):
        v = other.val if isinstance(other, Number) else other
        self.val += v
        return self

    def __repr__(self):
        return f"Num({self.val})"

n1 = Number(10)
n2 = Number(20)
__check(__line(n1 + n2), "Num(30)")
__check(__line(100 + n1), "Num(110)")
n1 += 5
__check(__line(n1), "Num(15)")
