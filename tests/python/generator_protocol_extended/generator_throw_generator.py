# vybe-test: python/generator_protocol_extended/generator_throw_generator
# origin: languages/python/tests/python/test_generator_protocol_extended.rs
# vybe-test-mode: compile

def g():
 yield 1
it = g()
it.throw(GeneratorExit)
