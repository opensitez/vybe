# vybe-test: python/new_features/os_mkdir
# origin: languages/python/tests/python/test_new_features.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import os
    os.mkdir('/tmp/testdir')

except BaseException:
    pass
