# vybe-test: python/oop_meta_descriptor_spec/abc_register_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

from abc import ABC
class Base(ABC):
    pass
Base.register(tuple)
