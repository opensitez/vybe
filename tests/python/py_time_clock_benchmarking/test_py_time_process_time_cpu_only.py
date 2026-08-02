# vybe-test: python/py_time_clock_benchmarking/test_py_time_process_time_cpu_only
# origin: languages/python/tests/python/test_py_time_clock_benchmarking.rs

import time

p1 = time.process_time()
# CPU-bound work
_ = [x * x for x in range(10000)]
p2 = time.process_time()

print(p2 >= p1)
