use super::helpers::run_python;

// ════════════════════════════════════════════════════════════
// Category: calendar module — month/year/weekday calculations
// ════════════════════════════════════════════════════════════

#[test]
fn test_calendar_isleap_known_years() {
    let out = run_python(
        r#"
import calendar
print(calendar.isleap(2000))
print(calendar.isleap(1900))
print(calendar.isleap(2024))
print(calendar.isleap(2023))
"#,
    );
    assert_eq!(out, vec!["True", "False", "True", "False"]);
}

#[test]
fn test_calendar_leapdays_range() {
    let out = run_python(
        r#"
import calendar
print(calendar.leapdays(2000, 2024))
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_calendar_monthrange_first_weekday() {
    let out = run_python(
        r#"
import calendar
# January 2024: Tuesday start (1), 31 days
wd, days = calendar.monthrange(2024, 1)
print(wd)
print(days)
"#,
    );
    assert_eq!(out, vec!["0", "31"]);
}

#[test]
fn test_calendar_weekday_function() {
    let out = run_python(
        r#"
import calendar
# Jan 1 2024 was a Monday
print(calendar.weekday(2024, 1, 1))
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_calendar_day_name_abbr() {
    let out = run_python(
        r#"
import calendar
print(list(calendar.day_name)[0])
print(list(calendar.day_abbr)[0])
"#,
    );
    assert_eq!(out, vec!["Monday", "Mon"]);
}

#[test]
fn test_calendar_month_name_abbr() {
    let out = run_python(
        r#"
import calendar
names = list(calendar.month_name)
abbrs = list(calendar.month_abbr)
print(names[1])
print(abbrs[1])
"#,
    );
    assert_eq!(out, vec!["January", "Jan"]);
}

#[test]
fn test_calendar_monthcalendar_shape() {
    let out = run_python(
        r#"
import calendar
mc = calendar.monthcalendar(2024, 2)
# February 2024 has 5 weeks
print(len(mc) >= 4)
print(len(mc[0]))
"#,
    );
    assert_eq!(out, vec!["True", "7"]);
}

#[test]
fn test_calendar_setfirstweekday_getfirstweekday() {
    let out = run_python(
        r#"
import calendar
calendar.setfirstweekday(6)  # Sunday
print(calendar.firstweekday())
calendar.setfirstweekday(0)  # reset to Monday
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_calendar_itermonthdates_count() {
    let out = run_python(
        r#"
import calendar
cal = calendar.Calendar()
dates = list(cal.itermonthdates(2024, 2))
print(len(dates) >= 28)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_calendar_itermonthdays2() {
    let out = run_python(
        r#"
import calendar
cal = calendar.Calendar(firstweekday=0)
days = [(d, wd) for d, wd in cal.itermonthdays2(2024, 1) if d != 0]
print(days[0])
print(len(days))
"#,
    );
    assert_eq!(out, vec!["(1, 0)", "31"]);
}

#[test]
fn test_calendar_yeardayscalendar_structure() {
    let out = run_python(
        r#"
import calendar
cal = calendar.Calendar()
ydc = cal.yeardayscalendar(2024, width=3)
print(len(ydc))  # 4 rows of 3 months each = 4
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_calendar_textcalendar_formatmonth() {
    let out = run_python(
        r#"
import calendar
tc = calendar.TextCalendar()
s = tc.formatmonth(2024, 1)
print("January 2024" in s)
print("Mo" in s or "Mon" in s)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_calendar_htmlcalendar_formatmonth() {
    let out = run_python(
        r#"
import calendar
hc = calendar.HTMLCalendar()
s = hc.formatmonth(2024, 1)
print("<table" in s.lower())
print("January" in s)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_calendar_month_constant() {
    let out = run_python(
        r#"
import calendar
print(calendar.MONDAY)
print(calendar.SUNDAY)
"#,
    );
    assert_eq!(out, vec!["0", "6"]);
}

#[test]
fn test_calendar_monthcalendar_zeros_padding() {
    let out = run_python(
        r#"
import calendar
mc = calendar.monthcalendar(2024, 1)
# First row may have 0s before the 1st
first_nonzero = next(d for row in mc for d in row if d != 0)
print(first_nonzero)
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_calendar_prmonth_no_crash() {
    let out = run_python(
        r#"
import calendar, io, sys
buf = io.StringIO()
sys.stdout = buf
calendar.prmonth(2024, 1)
sys.stdout = sys.__stdout__
print("January" in buf.getvalue())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_calendar_timegm_epoch() {
    let out = run_python(
        r#"
import calendar
# epoch: 1970-01-01 00:00:00 UTC
ts = calendar.timegm((1970, 1, 1, 0, 0, 0, 0, 1, 0))
print(ts)
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_calendar_itermonthdays_no_padding() {
    let out = run_python(
        r#"
import calendar
cal = calendar.Calendar()
days = [d for d in cal.itermonthdays(2024, 1) if d != 0]
print(len(days))
print(days[0])
print(days[-1])
"#,
    );
    assert_eq!(out, vec!["31", "1", "31"]);
}

#[test]
fn test_calendar_mdays_constant() {
    let out = run_python(
        r#"
import calendar
# mdays[0] is 0 (placeholder), mdays[1] is Jan=31
print(calendar.mdays[1])
print(calendar.mdays[2])
print(calendar.mdays[12])
"#,
    );
    assert_eq!(out, vec!["31", "28", "31"]);
}

#[test]
fn test_calendar_different_firstweekday_monthrange() {
    let out = run_python(
        r#"
import calendar
cal = calendar.Calendar(firstweekday=6)  # Sunday first
mc = cal.monthdayscalendar(2024, 1)
# With Sunday first, first col is Sunday
print(len(mc[0]) == 7)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
