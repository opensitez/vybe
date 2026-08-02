# vybe-test: python/oop_meta_descriptor_spec/mro_method_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class A: pass
class B(A): pass
order = B.mro()
