# vybe-test: python/py_string_formatting_template/test_py_custom_format_dunder
# origin: languages/python/tests/python/test_py_string_formatting_template.rs

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

class Money:
    def __init__(self, amount, currency="USD"):
        self.amount = amount
        self.currency = currency

    def __format__(self, format_spec):
        if format_spec == "code":
            return f"{self.amount:.2f} {self.currency}"
        elif format_spec == "symbol":
            sym = "$" if self.currency == "USD" else "€"
            return f"{sym}{self.amount:.2f}"
        return f"{self.amount:.2f}"

m = Money(49.5, "USD")
__check(__line(f"{m:code}"), "49.50 USD")
__check(__line(f"{m:symbol}"), "$49.50")
__check(__line(f"{m}"), "49.50")
