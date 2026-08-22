# vybe-test: python/oop_meta_descriptor_spec/super_in_staticmethod_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

class A:
    @staticmethod
    def f():
        return 1
class B(A):
    @staticmethod
    def g():
        return A.f()
