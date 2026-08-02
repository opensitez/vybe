# vybe-test: python/generator_iterator_protocol_spec/generator_throw_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def gen():
    try:
        yield 1
    except ValueError:
        yield 2

g = gen()
next(g)
g.throw(ValueError())
