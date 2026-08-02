# vybe-test: python/descriptor_metaclass_extended/slots_weakref
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs
# vybe-test-mode: compile

class C:
 __slots__ = ('__weakref__', 'x')
