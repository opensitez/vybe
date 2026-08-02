# vybe-test: python/py_getattr_slot/plain_miss_raises_attribute_error
# origin: languages/python/tests/python/test_py_getattr_slot.rs

class Plain:
    def __init__(self):
        self.here = 1

p = Plain()
try:
    print(p.nope)
except AttributeError:
    print("AttributeError")
except KeyError:
    print("KeyError")
