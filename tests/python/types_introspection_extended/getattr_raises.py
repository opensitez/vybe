# vybe-test: python/types_introspection_extended/getattr_raises
# origin: languages/python/tests/python/test_types_introspection_extended.rs
# vybe-test-mode: compile

class C: pass
try:
 getattr(C(), 'x')
except AttributeError:
 pass
