# vybe-test: python/oop_meta_descriptor_spec/subclasshook_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

from abc import ABCMeta
class Base(metaclass=ABCMeta):
    @classmethod
    def __subclasshook__(cls, C):
        return NotImplemented
