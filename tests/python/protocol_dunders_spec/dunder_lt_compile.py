# vybe-test: python/protocol_dunders_spec/dunder_lt_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Box:
    def __lt__(self, other):
        return True
