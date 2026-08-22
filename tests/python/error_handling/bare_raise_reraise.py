# vybe-test: python/error_handling/bare_raise_reraise
# origin: languages/python/tests/python/test_error_handling.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    try:
        x = 1 / 0
    except:
        print('logging')
        raise

except BaseException:
    pass
