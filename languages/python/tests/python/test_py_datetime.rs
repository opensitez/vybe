use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: datetime — date, time, datetime, timedelta, timezone, strftime, strptime, isoformat
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_datetime_date_creation_and_attributes() {
    let src = r#"
from datetime import date

d = date(2024, 3, 15)
print(d.year, d.month, d.day)
print(d.isoformat())
print(d.weekday())  # 0=Monday; 2024-03-15 is Friday=4
print(d.strftime("%d/%m/%Y"))
"#;
    assert_eq!(
        run_python(src),
        vec!["2024 3 15", "2024-03-15", "4", "15/03/2024"]
    );
}

#[test]
fn test_py_datetime_date_arithmetic() {
    let src = r#"
from datetime import date, timedelta

d = date(2024, 1, 1)
d2 = d + timedelta(days=100)
print(d2.isoformat())
diff = date(2024, 12, 31) - date(2024, 1, 1)
print(diff.days)
"#;
    assert_eq!(run_python(src), vec!["2024-04-10", "365"]);
}

#[test]
fn test_py_datetime_datetime_creation_and_now() {
    let src = r#"
from datetime import datetime

dt = datetime(2024, 6, 15, 10, 30, 45)
print(dt.year, dt.month, dt.day)
print(dt.hour, dt.minute, dt.second)
print(dt.isoformat())

now = datetime.now()
print(isinstance(now, datetime))
"#;
    assert_eq!(
        run_python(src),
        vec!["2024 6 15", "10 30 45", "2024-06-15T10:30:45", "True"]
    );
}

#[test]
fn test_py_datetime_timedelta_operations() {
    let src = r#"
from datetime import timedelta

td1 = timedelta(days=2, hours=3, minutes=30)
td2 = timedelta(hours=6)
print((td1 + td2).total_seconds())
print(td1.days)
print(td1.seconds)  # only the time portion in seconds
"#;
    assert_eq!(run_python(src), vec!["84600.0", "2", "12600"]);
}

#[test]
fn test_py_datetime_strptime_parsing() {
    let src = r#"
from datetime import datetime

dt = datetime.strptime("2024-03-15 10:30", "%Y-%m-%d %H:%M")
print(dt.year)
print(dt.month)
print(dt.day)
print(dt.hour)
"#;
    assert_eq!(run_python(src), vec!["2024", "3", "15", "10"]);
}

#[test]
fn test_py_datetime_timezone_aware() {
    let src = r#"
from datetime import datetime, timezone, timedelta

utc = timezone.utc
dt_utc = datetime(2024, 6, 15, 12, 0, tzinfo=utc)
print(dt_utc.isoformat())

ny_tz = timezone(timedelta(hours=-5))
dt_ny = dt_utc.astimezone(ny_tz)
print(dt_ny.hour)
print(dt_ny.utcoffset())
"#;
    assert_eq!(
        run_python(src),
        vec!["2024-06-15T12:00:00+00:00", "7", "-1 day, 19:00:00"]
    );
}

#[test]
fn test_py_datetime_comparison() {
    let src = r#"
from datetime import datetime

dt1 = datetime(2024, 1, 1)
dt2 = datetime(2024, 6, 15)
print(dt1 < dt2)
print(dt2 > dt1)
print(dt1 == datetime(2024, 1, 1))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_datetime_combine_date_and_time() {
    let src = r#"
from datetime import date, time, datetime

d = date(2024, 6, 15)
t = time(14, 30, 0)
dt = datetime.combine(d, t)
print(dt.isoformat())
print(dt.date() == d)
print(dt.time() == t)
"#;
    assert_eq!(run_python(src), vec!["2024-06-15T14:30:00", "True", "True"]);
}

#[test]
fn test_py_datetime_fromtimestamp_and_timestamp() {
    let src = r#"
from datetime import datetime, timezone

ts = 1704067200.0  # 2024-01-01 00:00:00 UTC
dt = datetime.fromtimestamp(ts, tz=timezone.utc)
print(dt.year, dt.month, dt.day)
print(dt.timestamp() == ts)
"#;
    assert_eq!(run_python(src), vec!["2024 1 1", "True"]);
}

#[test]
fn test_py_datetime_replace() {
    let src = r#"
from datetime import datetime

dt = datetime(2024, 6, 15, 10, 30)
dt2 = dt.replace(year=2025, hour=0, minute=0)
print(dt2.isoformat())
print(dt.isoformat())  # original unchanged
"#;
    assert_eq!(
        run_python(src),
        vec!["2025-06-15T00:00:00", "2024-06-15T10:30:00"]
    );
}

#[test]
fn test_py_datetime_fromisoformat() {
    let src = r#"
from datetime import datetime, date

d = date.fromisoformat("2024-03-15")
print(d.year, d.month)

dt = datetime.fromisoformat("2024-03-15T10:30:00")
print(dt.hour)
"#;
    assert_eq!(run_python(src), vec!["2024 3", "10"]);
}

#[test]
fn test_py_datetime_weekday_and_isocalendar() {
    let src = r#"
from datetime import date

d = date(2024, 1, 1)  # Monday
print(d.weekday())    # 0 = Monday
print(d.isoweekday()) # 1 = Monday
iso = d.isocalendar()
print(iso.year, iso.week, iso.weekday)
"#;
    assert_eq!(run_python(src), vec!["0", "1", "2024 1 1"]);
}

#[test]
fn test_py_datetime_timedelta_negative() {
    let src = r#"
from datetime import datetime, timedelta

dt = datetime(2024, 1, 15, 12, 0)
yesterday = dt - timedelta(days=1)
print(yesterday.day)

two_hours_ago = dt - timedelta(hours=2)
print(two_hours_ago.hour)
"#;
    assert_eq!(run_python(src), vec!["14", "10"]);
}

#[test]
fn test_py_datetime_min_max_resolution() {
    let src = r#"
from datetime import date, datetime, timedelta

print(date.min)
print(date.max)
print(date.resolution == timedelta(days=1))
print(datetime.resolution)
"#;
    assert_eq!(
        run_python(src),
        vec!["0001-01-01", "9999-12-31", "True", "0:00:00.000001"]
    );
}
