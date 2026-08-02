# vybe-test: python/protocol_dunders_spec/dunder_contains_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Bag:
    def __contains__(self, item):
        return False
