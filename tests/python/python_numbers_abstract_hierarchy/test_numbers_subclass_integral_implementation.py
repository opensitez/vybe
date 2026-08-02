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


def __check(got, want):
    if got != want:
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
__check(__line(isinstance(c, numbers.Integral)), "True")
__check(__line(int(c) == 10), "True")
__check(__line(c + 5 == 15), "True")
