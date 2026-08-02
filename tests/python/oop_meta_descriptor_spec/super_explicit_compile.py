# vybe-test: python/oop_meta_descriptor_spec/super_explicit_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class A:
    def f(self):
        return 1
class B(A):
    def f(self):
        return super(B, self).f()
