# vybe-test: python/protocol_dunders_spec/dunder_isub_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Counter:
    def __isub__(self, other):
        return self
