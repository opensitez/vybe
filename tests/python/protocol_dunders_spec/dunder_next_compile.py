# vybe-test: python/protocol_dunders_spec/dunder_next_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Seq:
    def __next__(self):
        raise StopIteration
