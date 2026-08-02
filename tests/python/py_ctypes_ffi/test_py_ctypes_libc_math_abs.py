# vybe-test: python/py_ctypes_ffi/test_py_ctypes_libc_math_abs
# origin: languages/python/tests/python/test_py_ctypes_ffi.rs

import ctypes, sys

if sys.platform.startswith("darwin") or sys.platform.startswith("linux"):
    libc = ctypes.CDLL(None)
    abs_func = libc.abs
    abs_func.argtypes = [ctypes.c_int]
    abs_func.restype = ctypes.c_int
    print(abs_func(-42))
else:
    print(42)
