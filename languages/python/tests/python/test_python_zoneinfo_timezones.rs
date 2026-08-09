use super::helpers::run_python;

// zoneinfo — ZoneInfo, available_timezones, reset_tzpath, ZoneInfoNotFoundError, datetime timezone conversions, fold handling

#[test]
fn test_zoneinfo_utc_timezone() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime
dt = datetime(2025, 1, 1, 12, 0, tzinfo=ZoneInfo("UTC"))
print(dt.tzname())
print(dt.utcoffset().total_seconds())
"#,
    );
    assert_eq!(out, vec!["UTC", "0.0"]);
}

#[test]
fn test_zoneinfo_available_timezones_nonempty() {
    let out = run_python(
        r#"
import zoneinfo
zones = zoneinfo.available_timezones()
print(isinstance(zones, set))
print("UTC" in zones)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_zoneinfo_invalid_timezone_raises_not_found_error() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
try:
    ZoneInfo("NonExistent/City_Location_XYZ")
except ZoneInfoNotFoundError:
    print("ZoneInfoNotFoundError")
"#,
    );
    assert_eq!(out, vec!["ZoneInfoNotFoundError"]);
}

#[test]
fn test_zoneinfo_datetime_astimezone_conversion() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime
utc_dt = datetime(2025, 6, 1, 12, 0, tzinfo=ZoneInfo("UTC"))
ny_dt = utc_dt.astimezone(ZoneInfo("America/New_York"))
print(ny_dt.hour)   # UTC 12:00 -> NY 08:00 (EDT = UTC-4)
print(ny_dt.tzname())
"#,
    );
    assert_eq!(out, vec!["8", "EDT"]);
}

#[test]
fn test_zoneinfo_key_attribute() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
zi = ZoneInfo("Europe/Paris")
print(zi.key)
"#,
    );
    assert_eq!(out, vec!["Europe/Paris"]);
}

#[test]
fn test_zoneinfo_str_and_repr() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
zi = ZoneInfo("Asia/Tokyo")
print(str(zi))
print(repr(zi))
"#,
    );
    assert_eq!(
        out,
        vec!["Asia/Tokyo", "zoneinfo.ZoneInfo(key='Asia/Tokyo')"]
    );
}

#[test]
fn test_zoneinfo_same_key_returns_same_instance() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
z1 = ZoneInfo("UTC")
z2 = ZoneInfo("UTC")
print(z1 is z2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zoneinfo_clear_cache() {
    let out = run_python(
        r#"
import zoneinfo
from zoneinfo import ZoneInfo
z1 = ZoneInfo("UTC")
zoneinfo.reset_tzpath()
z2 = ZoneInfo("UTC")
print(z1 == z2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zoneinfo_dst_transition_standard_time() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime
# Winter time in NY = EST (UTC-5)
dt = datetime(2025, 1, 15, 12, 0, tzinfo=ZoneInfo("America/New_York"))
print(dt.tzname())
print(dt.utcoffset().total_seconds() / 3600)
"#,
    );
    assert_eq!(out, vec!["EST", "-5.0"]);
}

#[test]
fn test_zoneinfo_dst_transition_daylight_time() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime
# Summer time in NY = EDT (UTC-4)
dt = datetime(2025, 7, 15, 12, 0, tzinfo=ZoneInfo("America/New_York"))
print(dt.tzname())
print(dt.utcoffset().total_seconds() / 3600)
"#,
    );
    assert_eq!(out, vec!["EDT", "-4.0"]);
}

#[test]
fn test_zoneinfo_fold_0_first_occurrence_ambiguous_wall_time() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime
# Fallback in NY: 1:30 AM occurs twice. fold=0 is first occurrence (EDT)
dt0 = datetime(2025, 11, 2, 1, 30, fold=0, tzinfo=ZoneInfo("America/New_York"))
print(dt0.tzname())
"#,
    );
    assert_eq!(out, vec!["EDT"]);
}

#[test]
fn test_zoneinfo_fold_1_second_occurrence_ambiguous_wall_time() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime
# Fallback in NY: fold=1 is second occurrence (EST)
dt1 = datetime(2025, 11, 2, 1, 30, fold=1, tzinfo=ZoneInfo("America/New_York"))
print(dt1.tzname())
"#,
    );
    assert_eq!(out, vec!["EST"]);
}

#[test]
fn test_zoneinfo_pickle_roundtrip() {
    let out = run_python(
        r#"
import pickle
from zoneinfo import ZoneInfo
zi = ZoneInfo("Pacific/Auckland")
data = pickle.dumps(zi)
restored = pickle.loads(data)
print(restored.key)
print(restored is zi)
"#,
    );
    assert_eq!(out, vec!["Pacific/Auckland", "True"]);
}

#[test]
fn test_zoneinfo_from_file_constructor() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
import os
# Check ZoneInfo.from_file works with explicit key
zi = ZoneInfo.no_cache("UTC")
print(zi.key)
"#,
    );
    assert_eq!(out, vec!["UTC"]);
}

#[test]
fn test_zoneinfo_tzpath_list() {
    let out = run_python(
        r#"
import zoneinfo
print(isinstance(zoneinfo.TZPATH, tuple))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zoneinfo_equality_comparison() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
z1 = ZoneInfo("UTC")
z2 = ZoneInfo("Europe/London")
print(z1 == ZoneInfo("UTC"))
print(z1 == z2)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_zoneinfo_hashability() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
d = {ZoneInfo("UTC"): "universal", ZoneInfo("America/Chicago"): "central"}
print(d[ZoneInfo("UTC")])
"#,
    );
    assert_eq!(out, vec!["universal"]);
}

#[test]
fn test_zoneinfo_utcoffset_is_timedelta() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime, timedelta
dt = datetime(2025, 5, 1, 10, 0, tzinfo=ZoneInfo("UTC"))
print(isinstance(dt.utcoffset(), timedelta))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zoneinfo_dst_is_timedelta_or_none() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime, timedelta
dt = datetime(2025, 7, 1, 10, 0, tzinfo=ZoneInfo("America/New_York"))
print(isinstance(dt.dst(), timedelta))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zoneinfo_isoformat_with_timezone() {
    let out = run_python(
        r#"
from zoneinfo import ZoneInfo
from datetime import datetime
dt = datetime(2025, 1, 1, 12, 0, tzinfo=ZoneInfo("UTC"))
print(dt.isoformat())
"#,
    );
    assert_eq!(out, vec!["2025-01-01T12:00:00+00:00"]);
}
