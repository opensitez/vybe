# vybe-test: python/oop_meta_descriptor_spec/multiple_inheritance_super_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

class A: pass
class B(A): pass
class C(A): pass
class D(B, C):
    def f(self):
        return super().f()
