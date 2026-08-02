# vybe-test: python/python_calendar_module/test_calendar_monthcalendar_zeros_padding
# origin: languages/python/tests/python/test_python_calendar_module.rs

import calendar
mc = calendar.monthcalendar(2024, 1)
# First row may have 0s before the 1st
first_nonzero = next(d for row in mc for d in row if d != 0)
print(first_nonzero)
