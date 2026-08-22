# vybe-test: python/oop_inheritance_extended/oop_abstract_base
# origin: languages/python/tests/python/test_oop_inheritance_extended.rs

from abc import ABC, abstractmethod
class B(ABC):
 @abstractmethod
 def m(self):
  pass
