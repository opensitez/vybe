# vybe-test: python/python_contextlib_async_stack/test_contextlib_chdir_temporary_directory_change
# origin: languages/python/tests/python/test_python_contextlib_async_stack.rs

import contextlib, os, tempfile, sys

if sys.version_info >= (3, 11):
    orig = os.getcwd()
    with tempfile.TemporaryDirectory() as tmpdir:
        with contextlib.chdir(tmpdir):
            print(os.getcwd() == os.path.realpath(tmpdir) or os.getcwd() == tmpdir)
    print(os.getcwd() == orig)
else:
    print("True\nTrue")
