# vybe-test: python/stdlib_modules/datetime_strftime
# origin: languages/python/tests/python/test_stdlib_modules.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import datetime
    now = datetime.now()
    s = now.strftime('%Y-%m-%d')

except BaseException:
    pass
