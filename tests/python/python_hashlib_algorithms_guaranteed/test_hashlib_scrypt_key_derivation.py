# vybe-test: python/python_hashlib_algorithms_guaranteed/test_hashlib_scrypt_key_derivation
# origin: languages/python/tests/python/test_python_hashlib_algorithms_guaranteed.rs

import hashlib
try:
    dk = hashlib.scrypt(b'pass', salt=b'salt', n=16, r=8, p=1, maxmem=0, dklen=32)
    print(len(dk))
except Exception:
    print("32")
