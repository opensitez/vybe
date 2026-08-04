# vybe-test: python/python_calendar_module/test_calendar_prmonth_no_crash
# origin: languages/python/tests/python/test_python_calendar_module.rs

import calendar, io, sys
buf = io.StringIO()
sys.stdout = buf
calendar.prmonth(2024, 1)
sys.stdout = sys.__stdout__
print("January" in buf.getvalue())
