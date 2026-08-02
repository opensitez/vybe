# vybe-test: python/py_security/test_py_hashlib_different_algorithms
# origin: languages/python/tests/python/test_py_security.rs

import hashlib

algos = ["sha256", "sha512", "sha1", "md5"]
for name in algos:
    h = hashlib.new(name, b"test")
    print(f"{name}: {h.digest_size} bytes, {len(h.hexdigest())} hex chars")
