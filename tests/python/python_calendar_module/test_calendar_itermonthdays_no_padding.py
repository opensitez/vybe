# vybe-test: python/python_calendar_module/test_calendar_itermonthdays_no_padding
# origin: languages/python/tests/python/test_python_calendar_module.rs

import calendar
cal = calendar.Calendar()
days = [d for d in cal.itermonthdays(2024, 1) if d != 0]
print(len(days))
print(days[0])
print(days[-1])
