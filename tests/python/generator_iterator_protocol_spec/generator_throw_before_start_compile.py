# vybe-test: python/generator_iterator_protocol_spec/generator_throw_before_start_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    def gen():
        try:
            yield 1
        except ValueError:
            yield 2

    g = gen()
    g.throw(ValueError())

except BaseException:
    pass
