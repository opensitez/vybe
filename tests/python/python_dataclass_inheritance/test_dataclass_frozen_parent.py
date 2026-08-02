# vybe-test: python/python_dataclass_inheritance/test_dataclass_frozen_parent
# origin: languages/python/tests/python/test_python_dataclass_inheritance.rs

from dataclasses import dataclass

@dataclass(frozen=True)
class Immutable:
    value: int

obj = Immutable(42)
print(obj.value)
try:
    obj.value = 99
    print("no_error")
except (AttributeError, TypeError, Exception):
    print("immutable")
