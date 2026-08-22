# vybe-test: python/stdlib_compile_extended/abc_abstract
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

from abc import ABC, abstractmethod
class B(ABC):
 @abstractmethod
 def m(self): pass
