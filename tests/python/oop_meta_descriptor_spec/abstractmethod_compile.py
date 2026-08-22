# vybe-test: python/oop_meta_descriptor_spec/abstractmethod_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

from abc import ABC, abstractmethod
class Base(ABC):
    @abstractmethod
    def run(self):
        pass
