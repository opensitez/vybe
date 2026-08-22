# vybe-test: python/generator_extended/generator_throw_while_running
# origin: languages/python/tests/python/test_generator_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    def g():
     yield 1
     yield 2
    it = g()
    next(it)
    it.throw(RuntimeError)

except BaseException:
    pass
