# vybe-test: python/python_typing_typeddict_required/test_typing_is_typeddict_check
# origin: languages/python/tests/python/test_python_typing_typeddict_required.rs

from typing import TypedDict, is_typeddict, sys

if sys.version_info >= (3, 10):
    class Car(TypedDict): make: str
    class Normal: pass
    print(is_typeddict(Car))
    print(is_typeddict(Normal))
else:
    print("True\nFalse")
