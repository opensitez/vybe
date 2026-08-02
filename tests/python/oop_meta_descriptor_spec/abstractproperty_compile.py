# vybe-test: python/oop_meta_descriptor_spec/abstractproperty_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

from abc import ABC, abstractmethod
class Base(ABC):
    @property
    @abstractmethod
    def value(self):
        pass
