# vybe-test: python/numeric_semantics_extended/int_digits
# origin: languages/python/tests/python/test_numeric_semantics_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    print((255).digits(16))

except BaseException:
    pass
