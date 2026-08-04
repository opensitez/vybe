# vybe-test: python/python_numbers_abstract_hierarchy/test_numbers_subclass_integral_implementation
# origin: languages/python/tests/python/test_python_numbers_abstract_hierarchy.rs

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


# Output is COLLECTED, not paired. The emitter rewrites every `print(a, b)`
# into `__p(__line(a, b))`, appending here, and compares the whole buffer once
# at the end of the file. Pairing the i-th print with the i-th expected line
# cannot assert anything about a loop — 936 of Python's cases.
#
# `__p`/`__pr` take an ALREADY-JOINED string, not `*args` plus keyword-only
# terminator parameters. Those are broken under Vybe (measured: with a
# keyword-only sep/end after *args, the call appends nothing at all), while the
# plain `__line(*args)` above works. So the newline decision is made by WHICH
# helper the emitter calls.
#
# A comment in the FIRST position of an indented block used to be a parse error
# under Vybe — `def f():` followed by a comment line. Fixed in
# `languages/python/src/grammar.pest`: the preprocessor emits a comment-only
# line without an INDENT marker but still emits its NEWLINE, so `block` has to
# accept `":" NEWLINE NEWLINE* … INDENT`. Mid-block comments always worked,
# which is why this hid for so long.
__buf = ""


def __p(s):
    global __buf
    __buf += s + "\n"


def __pr(s):
    global __buf
    __buf += s


def __check(got, want):
    # The final print contributes a trailing newline that the expected line
    # vector never carried, so both forms are accepted.
    if got != want and got != want + "\n":
        print("FAIL: want [" + want + "] got [" + got + "]")
        raise Exception("assertion failed")

import numbers

class CustomInt(numbers.Integral):
    def __init__(self, val):
        self.val = val
    def __int__(self):
        return int(self.val)
    def __index__(self):
        return int(self.val)
    def __abs__(self):
        return abs(self.val)
    def __add__(self, other):
        return self.val + other
    def __radd__(self, other):
        return other + self.val
    def __sub__(self, other):
        return self.val - other
    def __rsub__(self, other):
        return other - self.val
    def __mul__(self, other):
        return self.val * other
    def __rmul__(self, other):
        return other * self.val
    def __truediv__(self, other):
        return self.val / other
    def __rtruediv__(self, other):
        return other / self.val
    def __floordiv__(self, other):
        return self.val // other
    def __rfloordiv__(self, other):
        return other // self.val
    def __mod__(self, other):
        return self.val % other
    def __rmod__(self, other):
        return other % self.val
    def __pow__(self, other):
        return self.val ** other
    def __rpow__(self, other):
        return other ** self.val
    def __rshift__(self, other):
        return self.val >> other
    def __rrshift__(self, other):
        return other >> self.val
    def __lshift__(self, other):
        return self.val << other
    def __rlshift__(self, other):
        return other << self.val
    def __and__(self, other):
        return self.val & other
    def __rand__(self, other):
        return other & self.val
    def __or__(self, other):
        return self.val | other
    def __ror__(self, other):
        return other | self.val
    def __xor__(self, other):
        return self.val ^ other
    def __rxor__(self, other):
        return other ^ self.val
    def __neg__(self):
        return -self.val
    def __pos__(self):
        return +self.val
    def __invert__(self):
        return ~self.val
    def __eq__(self, other):
        return self.val == other
    def __lt__(self, other):
        return self.val < other
    def __le__(self, other):
        return self.val <= other
    def __trunc__(self):
        return self.val
    def __floor__(self):
        return self.val
    def __ceil__(self):
        return self.val
    def __round__(self, n=None):
        return round(self.val, n)

c = CustomInt(10)
__p(__line(isinstance(c, numbers.Integral)))
__p(__line(int(c) == 10))
__p(__line(c + 5 == 15))
__check(__buf, "True\nTrue\nTrue")
