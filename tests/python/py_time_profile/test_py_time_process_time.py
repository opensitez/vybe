# vybe-test: python/py_time_profile/test_py_time_process_time
# origin: languages/python/tests/python/test_py_time_profile.rs

import time

t0 = time.process_time()
_ = sum(i * i for i in range(10000))
t1 = time.process_time()
print(t1 >= t0)
