# vybe-test: python/py_descriptors/test_py_descriptor_shared_storage_for_multiple_instances
# origin: languages/python/tests/python/test_py_descriptors.rs

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

class SlotDescriptor:
    def __set_name__(self, owner, name):
        self.slot_key = f"_slot_{name}"

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return getattr(obj, self.slot_key, "default")

    def __set__(self, obj, value):
        setattr(obj, self.slot_key, value)

class Widget:
    color = SlotDescriptor()
    size = SlotDescriptor()

w1, w2 = Widget(), Widget()
w1.color = "red"
w2.color = "blue"
w1.size = "small"
__check(__line(w1.color, w2.color, w1.size, w2.size), "red blue small default")
