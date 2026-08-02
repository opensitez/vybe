# vybe-test: python/generator_iterator_protocol_spec/yield_in_with_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def gen():
    with open('x') as f:
        yield f.read()
