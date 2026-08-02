# vybe-test: python/python_hashlib_algorithms_guaranteed/test_hashlib_file_digest_helper
# origin: languages/python/tests/python/test_python_hashlib_algorithms_guaranteed.rs

import hashlib, io, sys
if sys.version_info >= (3, 11):
    f = io.BytesIO(b"file content")
    h = hashlib.file_digest(f, "sha256")
    print(h.hexdigest() == hashlib.sha256(b"file content").hexdigest())
else:
    print(True)
