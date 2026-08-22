# vybe-test: python/protocol_dunders_spec/dunder_floordiv_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Vec:
    def __floordiv__(self, other):
        return self
