# vybe-test: python/datetime_runtime/calendar_calendar_iter
# origin: languages/python/tests/python/test_datetime_runtime.rs
# vybe-test-mode: compile

import calendar
cal = calendar.Calendar()
list(cal.itermonthdays(2020, 6))
