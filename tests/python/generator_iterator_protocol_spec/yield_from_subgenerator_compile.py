# vybe-test: python/generator_iterator_protocol_spec/yield_from_subgenerator_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def sub():
    yield 1
    yield 2
def gen():
    yield from sub()
