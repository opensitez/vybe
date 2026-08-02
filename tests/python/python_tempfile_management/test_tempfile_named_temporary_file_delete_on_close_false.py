# vybe-test: python/python_tempfile_management/test_tempfile_named_temporary_file_delete_on_close_false
# origin: languages/python/tests/python/test_python_tempfile_management.rs

import tempfile, os, sys
if sys.version_info >= (3, 12):
    f = tempfile.NamedTemporaryFile(delete_on_close=False)
    name = f.name
    f.close()
    print(os.path.exists(name))
    os.unlink(name)
else:
    print(True)
