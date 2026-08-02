# vybe-test: python/datetime_runtime/datetime_astimezone
# origin: languages/python/tests/python/test_datetime_runtime.rs
# vybe-test-mode: compile

import datetime
dt = datetime.datetime(2020,1,1,tzinfo=datetime.timezone.utc)
dt.astimezone()
