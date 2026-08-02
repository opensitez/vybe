# vybe-test: python/py_shutil_tempfile/test_py_shutil_get_archive_formats
# origin: languages/python/tests/python/test_py_shutil_tempfile.rs

import shutil

formats = [name for name, _ in shutil.get_archive_formats()]
print("zip" in formats)
print("tar" in formats)
