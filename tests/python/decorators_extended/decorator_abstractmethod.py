# vybe-test: python/decorators_extended/decorator_abstractmethod
# origin: languages/python/tests/python/test_decorators_extended.rs

from abc import abstractmethod
class B:
 @abstractmethod
 def m(self):
  pass
