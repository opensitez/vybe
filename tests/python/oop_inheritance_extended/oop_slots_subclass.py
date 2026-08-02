# vybe-test: python/oop_inheritance_extended/oop_slots_subclass
# origin: languages/python/tests/python/test_oop_inheritance_extended.rs
# vybe-test-mode: compile

class B:
 __slots__ = ()
class D(B):
 __slots__ = ('x',)
