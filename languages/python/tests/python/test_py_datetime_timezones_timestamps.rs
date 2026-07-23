use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Datetime, Timezones & Timestamps — datetime, date, time, timedelta, timezone.utc, ISO 8601, timestamp conversion
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_datetime_arithmetic_timedelta() {
    let src = r#"
from datetime import datetime, timedelta

dt = datetime(2024, 1, 1, 12, 0, 0)
dt_future = dt + timedelta(days=10, hours=5)
print(dt_future)

diff = dt_future - dt
print(diff.days, diff.seconds)
"#;
    assert_eq!(run_python(src), vec!["2024-01-11 17:00:00", "10 18000"]);
}

#[test]
fn test_py_datetime_isoformat_fromisoformat() {
    let src = r#"
from datetime import datetime

dt = datetime(2024, 5, 12, 15, 30, 45)
iso_str = dt.isoformat()
print(iso_str)

dt_parsed = datetime.fromisoformat(iso_str)
print(dt_parsed == dt)
"#;
    assert_eq!(run_python(src), vec!["2024-05-12T15:30:45", "True"]);
}

#[test]
fn test_py_datetime_utc_timezone_awareness() {
    let src = r#"
from datetime import datetime, timezone, timedelta

utc_now = datetime(2024, 6, 1, 12, 0, 0, tzinfo=timezone.utc)
print(utc_now.tzname())

custom_tz = timezone(timedelta(hours=5, minutes=30))
local_dt = utc_now.astimezone(custom_tz)
print(local_dt.strftime("%Y-%m-%d %H:%M:%S %z"))
"#;
    assert_eq!(run_python(src), vec!["UTC", "2024-06-01 17:30:00 +0530"]);
}

#[test]
fn test_py_datetime_timestamp_conversions() {
    let src = r#"
from datetime import datetime, timezone

dt = datetime(2024, 1, 1, 0, 0, 0, tzinfo=timezone.utc)
ts = dt.timestamp()
print(int(ts))

dt_back = datetime.fromtimestamp(ts, tz=timezone.utc)
print(dt_back == dt)
"#;
    assert_eq!(run_python(src), vec!["1704067200", "True"]);
}

#[test]
fn test_py_date_today_weekday_isoweekday() {
    let src = r#"
from datetime import date

d = date(2024, 5, 12)  # Sunday
print(d.year, d.month, d.day)
print(d.weekday())     # Monday=0 ... Sunday=6
print(d.isoweekday())  # Monday=1 ... Sunday=7
"#;
    assert_eq!(run_python(src), vec!["2024 5 12", "6", "7"]);
}

#[test]
fn test_py_time_components_and_formatting() {
    let src = r#"
from datetime import time

t = time(14, 30, 45, 123456)
print(t.hour, t.minute, t.second, t.microsecond)
print(t.strftime("%H:%M:%S.%f"))
"#;
    assert_eq!(run_python(src), vec!["14 30 45 123456", "14:30:45.123456"]);
}

#[test]
fn test_py_datetime_combine_date_and_time() {
    let src = r#"
from datetime import date, time, datetime

d = date(2024, 10, 31)
t = time(23, 59, 59)
dt = datetime.combine(d, t)
print(dt)
"#;
    assert_eq!(run_python(src), vec!["2024-10-31 23:59:59"]);
}

#[test]
fn test_py_datetime_strptime_parsing() {
    let src = r#"
from datetime import datetime

date_str = "12/May/2024:15:30:45 +0000"
dt = datetime.strptime(date_str, "%d/%b/%Y:%H:%M:%S %z")
print(dt.year, dt.month, dt.day)
"#;
    assert_eq!(run_python(src), vec!["2024 5 12"]);
}

#[test]
fn test_py_timedelta_total_seconds() {
    let src = r#"
from datetime import timedelta

td = timedelta(days=2, hours=3, minutes=30)
print(td.total_seconds())
"#;
    assert_eq!(run_python(src), vec!["185400.0"]);
}

#[test]
fn test_py_datetime_replace_components() {
    let src = r#"
from datetime import datetime

dt = datetime(2024, 1, 1, 10, 0, 0)
dt_modified = dt.replace(year=2025, month=12, day=25)
print(dt_modified)
"#;
    assert_eq!(run_python(src), vec!["2025-12-25 10:00:00"]);
}
