# vybe-test: python/oop_inheritance_extended/oop_metaclass_type
# origin: languages/python/tests/python/test_oop_inheritance_extended.rs
# vybe-test-mode: compile

class M(type):
 pass
class C(metaclass=M):
 pass
