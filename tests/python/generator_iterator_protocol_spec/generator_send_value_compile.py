# vybe-test: python/generator_iterator_protocol_spec/generator_send_value_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def gen():
    x = yield 1
    yield x

g = gen()
next(g)
g.send(42)
