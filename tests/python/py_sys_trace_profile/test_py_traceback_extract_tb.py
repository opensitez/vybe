# vybe-test: python/py_sys_trace_profile/test_py_traceback_extract_tb
# origin: languages/python/tests/python/test_py_sys_trace_profile.rs

import traceback, sys

def func_a():
    func_b()

def func_b():
    raise RuntimeError("error in b")

try:
    func_a()
except RuntimeError as e:
    tb = e.__traceback__
    frames = traceback.extract_tb(tb)
    func_names = [f.name for f in frames]
    print("func_a" in func_names)
    print("func_b" in func_names)
