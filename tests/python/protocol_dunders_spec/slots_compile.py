# vybe-test: python/protocol_dunders_spec/slots_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Point:
    __slots__ = ('x', 'y')
