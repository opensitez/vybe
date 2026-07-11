//! datetime, date, time, timedelta, timezone, zoneinfo.

crate::runtime_case!(
    datetime_date_year,
    "import datetime\nprint(datetime.date(2020, 6, 15).year)\n",
    "2020"
);
crate::runtime_case!(
    datetime_date_month,
    "import datetime\nprint(datetime.date(2020, 6, 15).month)\n",
    "6"
);
crate::runtime_case!(
    datetime_date_day,
    "import datetime\nprint(datetime.date(2020, 6, 15).day)\n",
    "15"
);
crate::runtime_case!(
    datetime_date_isoformat,
    "import datetime\nprint(datetime.date(2020, 6, 15).isoformat())\n",
    "2020-06-15"
);
crate::runtime_case!(
    datetime_date_weekday,
    "import datetime\nprint(datetime.date(2020, 6, 15).weekday())\n",
    "0"
);
crate::runtime_case!(
    datetime_date_replace,
    "import datetime\nprint(datetime.date(2020, 1, 1).replace(year=2021).year)\n",
    "2021"
);
crate::runtime_case!(
    datetime_time_hour,
    "import datetime\nprint(datetime.time(12, 30, 45).hour)\n",
    "12"
);
crate::runtime_case!(
    datetime_time_minute,
    "import datetime\nprint(datetime.time(12, 30, 45).minute)\n",
    "30"
);
crate::runtime_case!(
    datetime_time_second,
    "import datetime\nprint(datetime.time(12, 30, 45).second)\n",
    "45"
);
crate::runtime_case!(
    datetime_time_isoformat,
    "import datetime\nprint(datetime.time(12, 30, 45).isoformat())\n",
    "12:30:45"
);
crate::runtime_case!(
    datetime_datetime_construct,
    "import datetime\ndt = datetime.datetime(2020, 6, 15, 10, 30)\nprint(dt.year, dt.month)\n",
    "2020 6"
);
crate::runtime_case!(
    datetime_datetime_isoformat,
    "import datetime\nprint(datetime.datetime(2020, 6, 15, 10, 30).isoformat())\n",
    "2020-06-15T10:30:00"
);
crate::runtime_case!(
    datetime_datetime_timestamp,
    "import datetime\ndt = datetime.datetime(1970, 1, 1)\nprint(int(dt.timestamp()) >= 0)\n",
    "True"
);
crate::runtime_case!(
    datetime_datetime_fromtimestamp,
    "import datetime\ndt = datetime.datetime.fromtimestamp(0)\nprint(dt.year)\n",
    "1970"
);
crate::runtime_case!(
    datetime_datetime_combine,
    "import datetime\nd = datetime.date(2020, 1, 1)\nt = datetime.time(12, 0)\nprint(datetime.datetime.combine(d, t).hour)\n",
    "12"
);
crate::runtime_case!(
    datetime_timedelta_days,
    "import datetime\nprint(datetime.timedelta(days=2).days)\n",
    "2"
);
crate::runtime_case!(
    datetime_timedelta_seconds,
    "import datetime\nprint(datetime.timedelta(seconds=90).seconds)\n",
    "90"
);
crate::runtime_case!(
    datetime_timedelta_total_seconds,
    "import datetime\nprint(datetime.timedelta(hours=1).total_seconds())\n",
    "3600.0"
);
crate::runtime_case!(
    datetime_date_add_timedelta,
    "import datetime\nd = datetime.date(2020, 1, 1)\nprint((d + datetime.timedelta(days=1)).day)\n",
    "2"
);
crate::runtime_case!(
    datetime_datetime_subtract,
    "import datetime\na = datetime.datetime(2020, 1, 2)\nb = datetime.datetime(2020, 1, 1)\nprint((a - b).days)\n",
    "1"
);
crate::runtime_case!(
    datetime_date_comparison,
    "import datetime\nprint(datetime.date(2020, 1, 2) > datetime.date(2020, 1, 1))\n",
    "True"
);
crate::runtime_case!(
    datetime_min_max,
    "import datetime\nprint(datetime.date.min.year)\n",
    "1"
);
crate::runtime_case!(
    datetime_timezone_utc,
    "import datetime\nprint(datetime.timezone.utc.utcoffset(None).total_seconds())\n",
    "0.0"
);
crate::runtime_case!(
    datetime_timezone_fixed_offset,
    "import datetime\ntz = datetime.timezone(datetime.timedelta(hours=2))\nprint(tz.utcoffset(None).total_seconds())\n",
    "7200.0"
);
crate::runtime_case!(
    datetime_strptime,
    "import datetime\ndt = datetime.datetime.strptime('2020-06-15', '%Y-%m-%d')\nprint(dt.month)\n",
    "6"
);
crate::runtime_case!(
    datetime_strftime,
    "import datetime\ndt = datetime.datetime(2020, 6, 15)\nprint(dt.strftime('%Y'))\n",
    "2020"
);
crate::runtime_case!(
    datetime_date_fromisoformat,
    "import datetime\nprint(datetime.date.fromisoformat('2020-06-15').day)\n",
    "15"
);
crate::runtime_case!(
    datetime_time_fromisoformat,
    "import datetime\nprint(datetime.time.fromisoformat('12:30:45').minute)\n",
    "30"
);
crate::runtime_case!(
    datetime_datetime_fromisoformat,
    "import datetime\nprint(datetime.datetime.fromisoformat('2020-06-15T10:30:00').hour)\n",
    "10"
);
crate::runtime_case!(
    zoneinfo_utc,
    "from zoneinfo import ZoneInfo\nprint(ZoneInfo('UTC').key)\n",
    "UTC"
);
crate::runtime_case!(
    zoneinfo_list,
    "from zoneinfo import available_timezones\nprint('UTC' in available_timezones())\n",
    "True"
);
crate::runtime_case!(
    calendar_monthrange,
    "import calendar\nprint(calendar.monthrange(2020, 2)[1])\n",
    "29"
);
crate::runtime_case!(
    calendar_isleap,
    "import calendar\nprint(calendar.isleap(2020))\n",
    "True"
);
crate::runtime_case!(
    calendar_weekday,
    "import calendar\nprint(calendar.weekday(2020, 6, 15))\n",
    "0"
);
crate::runtime_case!(
    calendar_month_name,
    "import calendar\nprint(calendar.month_name[6])\n",
    "June"
);
crate::runtime_case!(
    calendar_day_name,
    "import calendar\nprint(calendar.day_name[0])\n",
    "Monday"
);
crate::runtime_case!(
    datetime_date_timetuple,
    "import datetime\nprint(datetime.date(2020, 6, 15).timetuple().tm_mon)\n",
    "6"
);
crate::runtime_case!(
    datetime_datetime_date_method,
    "import datetime\ndt = datetime.datetime(2020, 6, 15, 10)\nprint(dt.date().day)\n",
    "15"
);
crate::runtime_case!(
    datetime_datetime_time_method,
    "import datetime\ndt = datetime.datetime(2020, 6, 15, 10, 30)\nprint(dt.time().hour)\n",
    "10"
);
crate::runtime_case!(
    datetime_timedelta_mul,
    "import datetime\nprint((datetime.timedelta(days=1) * 2).days)\n",
    "2"
);
crate::runtime_case!(
    datetime_timedelta_neg,
    "import datetime\nprint((-datetime.timedelta(days=1)).days)\n",
    "-1"
);
crate::runtime_case!(
    datetime_date_toordinal,
    "import datetime\nprint(datetime.date(2020, 1, 1).toordinal() > 0)\n",
    "True"
);
crate::runtime_case!(
    datetime_date_fromordinal,
    "import datetime\nprint(datetime.date.fromordinal(1).year)\n",
    "1"
);
crate::runtime_case!(
    datetime_resolution,
    "import datetime\nprint(datetime.timedelta.resolution.total_seconds() > 0)\n",
    "True"
);
crate::runtime_case!(
    datetime_now_callable,
    "import datetime\nprint(isinstance(datetime.datetime.now(), datetime.datetime))\n",
    "True"
);
crate::runtime_case!(
    datetime_utcnow_callable,
    "import datetime\nprint(isinstance(datetime.datetime.utcnow(), datetime.datetime))\n",
    "True"
);
crate::runtime_case!(
    datetime_today_callable,
    "import datetime\nprint(isinstance(datetime.date.today(), datetime.date))\n",
    "True"
);

crate::compile_case!(
    datetime_astimezone,
    "import datetime\ndt = datetime.datetime(2020,1,1,tzinfo=datetime.timezone.utc)\ndt.astimezone()\n"
);
crate::compile_case!(
    zoneinfo_datetime,
    "from zoneinfo import ZoneInfo\nimport datetime\ndt = datetime.datetime(2020,1,1,tzinfo=ZoneInfo('UTC'))\n"
);
crate::compile_case!(
    calendar_monthcalendar,
    "import calendar\ncalendar.monthcalendar(2020, 6)\n"
);
crate::compile_case!(
    calendar_calendar_iter,
    "import calendar\ncal = calendar.Calendar()\nlist(cal.itermonthdays(2020, 6))\n"
);
crate::compile_case!(
    datetime_fold_attribute,
    "import datetime\ndt = datetime.datetime(2020,1,1)\nhasattr(dt, 'fold')\n"
);
