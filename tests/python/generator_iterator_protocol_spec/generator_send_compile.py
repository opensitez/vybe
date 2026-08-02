# vybe-test: python/generator_iterator_protocol_spec/generator_send_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def gen():
    value = yield 1
    yield value

g = gen()
g.send(None)
