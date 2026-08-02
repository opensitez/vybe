# vybe-test: python/descriptor_metaclass_extended/metaclass_conflict
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs
# vybe-test-mode: compile

class M1(type): pass
class M2(type): pass
try:
 class C(metaclass=M1, metaclass=M2): pass
except TypeError: pass
