# vybe-test: python/generator_iterator_protocol_spec/iter_custom_for_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

class Counter:
    def __iter__(self):
        return self
    def __next__(self):
        raise StopIteration
for item in Counter():
    print(item)
