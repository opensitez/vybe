# vybe-test: python/time_extended/time_tzset_call
# origin: languages/python/tests/python/test_time_extended.rs

import time
try:
 time.tzset()
except AttributeError:
 pass
