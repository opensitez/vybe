# vybe-test: python/descriptor_metaclass_extended/abc_register_virtual
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs
# vybe-test-mode: compile

from abc import ABC
class B(ABC):
 pass
class C: pass
B.register(C)
