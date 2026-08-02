# vybe-test: python/python_tempfile_management/test_tempfile_temporary_directory_ignore_cleanup_errors
# origin: languages/python/tests/python/test_python_tempfile_management.rs

import tempfile, os, sys
if sys.version_info >= (3, 10):
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmpdir:
        pass
    print("ok")
else:
    print("ok")
