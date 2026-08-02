# vybe-test: python/time_module/time_sleep_inside_while_once
# origin: languages/python/tests/python/test_time_module.rs

import time
n = 0
while n < 1:
 time.sleep(0)
 n += 1
print(n)
