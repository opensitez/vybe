# vybe-test: python/python_dataclass_kwonly_slots/test_dataclass_slots_attribute
# origin: languages/python/tests/python/test_python_dataclass_kwonly_slots.rs

from dataclasses import dataclass
import sys

if sys.version_info >= (3, 10):
    @dataclass(slots=True)
    class User:
        name: str
        age: int

    u = User("Alice", 30)
    print(u.name)
    try:
        u.extra_field = "dynamic"
    except AttributeError:
        print("AttributeError")
else:
    print("Alice\nAttributeError")
