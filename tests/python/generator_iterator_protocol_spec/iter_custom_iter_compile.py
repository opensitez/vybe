# vybe-test: python/generator_iterator_protocol_spec/iter_custom_iter_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs

class Counter:
    def __iter__(self):
        return self
    def __next__(self):
        raise StopIteration
