# vybe-test: python/protocol_dunders_spec/dunder_iadd_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Counter:
    def __iadd__(self, other):
        return self
