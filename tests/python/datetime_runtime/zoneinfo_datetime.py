# vybe-test: python/datetime_runtime/zoneinfo_datetime
# origin: languages/python/tests/python/test_datetime_runtime.rs

from zoneinfo import ZoneInfo
import datetime
dt = datetime.datetime(2020,1,1,tzinfo=ZoneInfo('UTC'))
