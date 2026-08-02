# vybe-test: python/protocol_dunders_spec/dunder_iter_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Seq:
    def __iter__(self):
        return self
