# vybe-test: python/exceptions_extended/raise_not_implemented
# origin: languages/python/tests/python/test_exceptions_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    raise NotImplementedError('todo')

except BaseException:
    pass
