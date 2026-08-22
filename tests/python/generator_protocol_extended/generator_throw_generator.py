# vybe-test: python/generator_protocol_extended/generator_throw_generator
# origin: languages/python/tests/python/test_generator_protocol_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    def g():
     yield 1
    it = g()
    it.throw(GeneratorExit)

except BaseException:
    pass
