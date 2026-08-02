# vybe-test: python/protocol_dunders_spec/dunder_imul_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Counter:
    def __imul__(self, other):
        return self
