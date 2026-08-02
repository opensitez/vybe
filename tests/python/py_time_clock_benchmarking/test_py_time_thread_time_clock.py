# vybe-test: python/py_time_clock_benchmarking/test_py_time_thread_time_clock
# origin: languages/python/tests/python/test_py_time_clock_benchmarking.rs

import time

if hasattr(time, "thread_time"):
    tt = time.thread_time()
    print(isinstance(tt, float))
else:
    print("True")
