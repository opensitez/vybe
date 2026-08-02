# vybe-test: python/stdlib_modules/datetime_strftime
# origin: languages/python/tests/python/test_stdlib_modules.rs
# vybe-test-mode: compile

import datetime
now = datetime.now()
s = now.strftime('%Y-%m-%d')
