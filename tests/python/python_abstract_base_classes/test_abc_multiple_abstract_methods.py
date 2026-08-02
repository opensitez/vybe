# vybe-test: python/python_abstract_base_classes/test_abc_multiple_abstract_methods
# origin: languages/python/tests/python/test_python_abstract_base_classes.rs

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

from abc import ABC, abstractmethod

class Vehicle(ABC):
    @abstractmethod
    def fuel_type(self):
        pass
    @abstractmethod
    def max_speed(self):
        pass

class Car(Vehicle):
    def fuel_type(self):
        return "petrol"
    def max_speed(self):
        return 200

c = Car()
__check(__line(c.fuel_type()), "petrol")
__check(__line(c.max_speed()), "200")
