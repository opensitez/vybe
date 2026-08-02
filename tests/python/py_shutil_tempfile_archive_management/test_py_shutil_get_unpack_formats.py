# vybe-test: python/py_shutil_tempfile_archive_management/test_py_shutil_get_unpack_formats
# origin: languages/python/tests/python/test_py_shutil_tempfile_archive_management.rs

import shutil

formats = [fmt for fmt, _, _ in shutil.get_unpack_formats()]
print("zip" in formats)
print("tar" in formats)
