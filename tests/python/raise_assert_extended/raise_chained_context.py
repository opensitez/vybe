# vybe-test: python/raise_assert_extended/raise_chained_context
# origin: languages/python/tests/python/test_raise_assert_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    try:
     1/0
    except ZeroDivisionError as e:
     raise ValueError() from e

except BaseException:
    pass
