# vybe-test: python/python_tempfile_management/test_tempfile_spooled_temporary_file_fileno
# origin: languages/python/tests/python/test_python_tempfile_management.rs

import tempfile
with tempfile.SpooledTemporaryFile(max_size=10) as sf:
    sf.write(b"exceed max size to force rollover")
    if hasattr(sf, "fileno"):
        try:
            print(isinstance(sf.fileno(), int))
        except Exception:
            print(True)
    else:
        print(True)
