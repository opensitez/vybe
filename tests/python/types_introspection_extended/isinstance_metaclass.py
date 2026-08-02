# vybe-test: python/types_introspection_extended/isinstance_metaclass
# origin: languages/python/tests/python/test_types_introspection_extended.rs
# vybe-test-mode: compile

class M(type):
 pass
class C(metaclass=M):
 pass
isinstance(C(), C)
