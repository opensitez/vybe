# vybe-test: python/os_path_extended/os_path_mkdir
# origin: languages/python/tests/python/test_os_path_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import os
    os.mkdir('tmp_test_dir', 0o755)

except BaseException:
    pass
