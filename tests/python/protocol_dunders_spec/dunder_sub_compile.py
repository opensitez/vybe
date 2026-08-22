# vybe-test: python/protocol_dunders_spec/dunder_sub_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Vec:
    def __sub__(self, other):
        return self
