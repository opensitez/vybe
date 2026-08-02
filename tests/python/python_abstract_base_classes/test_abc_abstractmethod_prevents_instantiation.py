# vybe-test: python/python_abstract_base_classes/test_abc_abstractmethod_prevents_instantiation
# origin: languages/python/tests/python/test_python_abstract_base_classes.rs

from abc import ABC, abstractmethod

class Shape(ABC):
    @abstractmethod
    def area(self):
        pass

try:
    s = Shape()
    print("no_error")
except TypeError:
    print("TypeError")
