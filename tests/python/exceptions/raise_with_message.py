# vybe-test: python/exceptions/raise_with_message
# origin: languages/python/tests/python/test_exceptions.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    raise ValueError("bad input")

except BaseException:
    pass
