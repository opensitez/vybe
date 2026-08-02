# vybe-test: python/python_abstract_base_classes/test_abc_partial_implementation_still_abstract
# origin: languages/python/tests/python/test_python_abstract_base_classes.rs

from abc import ABC, abstractmethod

class Base(ABC):
    @abstractmethod
    def a(self):
        pass
    @abstractmethod
    def b(self):
        pass

class Partial(Base):
    def a(self):
        return 1

try:
    Partial()
    print("no_error")
except TypeError:
    print("TypeError_still_abstract")
