# vybe-test: python/exception_flow_spec/except_tuple_custom_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

class A(Exception): pass
class B(Exception): pass
try:
    risky()
except (A, B):
    pass
