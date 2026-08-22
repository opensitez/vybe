# vybe-test: python/exceptions/raise_from
# origin: languages/python/tests/python/test_exceptions.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    try:
        x = 1 / 0
    except ZeroDivisionError as e:
        raise ValueError("invalid") from e

except BaseException:
    pass
