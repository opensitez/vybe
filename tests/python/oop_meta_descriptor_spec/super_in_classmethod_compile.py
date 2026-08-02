# vybe-test: python/oop_meta_descriptor_spec/super_in_classmethod_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class A:
    @classmethod
    def f(cls):
        return 1
class B(A):
    @classmethod
    def f(cls):
        return super().f()
