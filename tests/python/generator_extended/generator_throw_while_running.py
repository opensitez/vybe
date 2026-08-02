# vybe-test: python/generator_extended/generator_throw_while_running
# origin: languages/python/tests/python/test_generator_extended.rs
# vybe-test-mode: compile

def g():
 yield 1
 yield 2
it = g()
next(it)
it.throw(RuntimeError)
