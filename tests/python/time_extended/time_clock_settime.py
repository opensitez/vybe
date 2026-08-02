# vybe-test: python/time_extended/time_clock_settime
# origin: languages/python/tests/python/test_time_extended.rs
# vybe-test-mode: compile

import time
hasattr(time, 'clock_settime')
