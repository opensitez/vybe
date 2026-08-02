# vybe-test: python/protocol_dunders_spec/dunder_invert_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Mask:
    def __invert__(self):
        return self
