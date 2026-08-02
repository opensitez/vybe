# vybe-test: python/py_sys_trace_profile_traceback/test_py_traceback_extract_tb_frame_stack
# origin: languages/python/tests/python/test_py_sys_trace_profile_traceback.rs

import traceback

def alpha():
    beta()

def beta():
    raise RuntimeError("error in beta")

try:
    alpha()
except RuntimeError as e:
    frames = traceback.extract_tb(e.__traceback__)
    func_names = [f.name for f in frames]
    print("alpha" in func_names)
    print("beta" in func_names)
