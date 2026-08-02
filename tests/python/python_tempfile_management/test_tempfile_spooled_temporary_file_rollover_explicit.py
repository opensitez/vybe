# vybe-test: python/python_tempfile_management/test_tempfile_spooled_temporary_file_rollover_explicit
# origin: languages/python/tests/python/test_python_tempfile_management.rs

import tempfile
with tempfile.SpooledTemporaryFile(max_size=100) as sf:
    sf.write(b"data")
    if hasattr(sf, "rollover"):
        sf.rollover()
        print(sf._rolled if hasattr(sf, "_rolled") else True)
    else:
        print(True)
