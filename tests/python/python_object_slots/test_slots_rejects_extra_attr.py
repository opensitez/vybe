# vybe-test: python/python_object_slots/test_slots_rejects_extra_attr
# origin: languages/python/tests/python/test_python_object_slots.rs

class Locked:
    __slots__ = ('x',)

obj = Locked()
obj.x = 10
try:
    obj.y = 20
    print("allowed")
except AttributeError:
    print("denied")
