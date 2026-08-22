# vybe-test: python/generator_iterator_protocol_spec/stop_iteration_custom_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs

class Done(Exception):
    pass
class I:
    def __iter__(self):
        return self
    def __next__(self):
        raise StopIteration
