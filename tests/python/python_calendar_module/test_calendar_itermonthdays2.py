# vybe-test: python/python_calendar_module/test_calendar_itermonthdays2
# origin: languages/python/tests/python/test_python_calendar_module.rs

import calendar
cal = calendar.Calendar(firstweekday=0)
days = [(d, wd) for d, wd in cal.itermonthdays2(2024, 1) if d != 0]
print(days[0])
print(len(days))
