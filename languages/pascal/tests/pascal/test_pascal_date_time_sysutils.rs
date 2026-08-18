use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 72: Date & Time Utilities (TDateTime & SysUtils Date Routines)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_datetime_encode_decode_date() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime; y, m, d: Word;
begin
  dt := EncodeDate(2026, 7, 24);
  DecodeDate(dt, y, m, d);
  WriteLn(y.ToString + '-' + m.ToString + '-' + d.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["2026-7-24"]);
}

#[test]
fn test_datetime_encode_decode_time() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime; h, m, s, ms: Word;
begin
  dt := EncodeTime(14, 30, 45, 500);
  DecodeTime(dt, h, m, s, ms);
  WriteLn(h.ToString + ':' + m.ToString + ':' + s.ToString + '.' + ms.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["14:30:45.500"]);
}

#[test]
fn test_datetime_formatdatetime() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime;
begin
  dt := EncodeDate(2025, 12, 31);
  WriteLn(FormatDateTime('yyyy-mm-dd', dt));
end.
"#,
    );
    assert_eq!(out, vec!["2025-12-31"]);
}

#[test]
fn test_datetime_isleapyear() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(IsLeapYear(2024));
  WriteLn(IsLeapYear(2025));
  WriteLn(IsLeapYear(2000));
  WriteLn(IsLeapYear(1900));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False", "True", "False"]);
}

#[test]
fn test_datetime_dayofweek() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime;
begin
  // Sunday = 1, Monday = 2, ...
  dt := EncodeDate(2026, 7, 26); // Sunday
  WriteLn(DayOfWeek(dt));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_datetime_incmonth() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime; y, m, d: Word;
begin
  dt := EncodeDate(2026, 1, 15);
  dt := IncMonth(dt, 2);
  DecodeDate(dt, y, m, d);
  WriteLn(m);
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_datetime_incday_via_addition() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime; y, m, d: Word;
begin
  dt := EncodeDate(2026, 5, 1);
  dt := dt + 10;
  DecodeDate(dt, y, m, d);
  WriteLn(d);
end.
"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_datetime_daysinmonth() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils;
begin
  WriteLn(DaysInMonth(EncodeDate(2024, 2, 1))); // Leap year
  WriteLn(DaysInMonth(EncodeDate(2025, 2, 1))); // Non-leap
end.
"#,
    );
    assert_eq!(out, vec!["29", "28"]);
}

#[test]
fn test_datetime_daysbetween() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var d1, d2: TDateTime;
begin
  d1 := EncodeDate(2026, 1, 1);
  d2 := EncodeDate(2026, 1, 11);
  WriteLn(DaysBetween(d2, d1));
end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_datetime_datetimetounix_unixtodatetime() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var dt1, dt2: TDateTime; unix: Int64;
begin
  dt1 := EncodeDateTime(2026, 1, 1, 0, 0, 0, 0);
  unix := DateTimeToUnix(dt1);
  dt2 := UnixToDateTime(unix);
  WriteLn(SameDateTime(dt1, dt2));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_datetime_timestamp_conversion() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt, dt2: TDateTime; ts: TTimeStamp;
begin
  dt := EncodeDate(2026, 3, 15);
  ts := DateTimeToTimeStamp(dt);
  dt2 := TimeStampToDateTime(ts);
  WriteLn(dt = dt2);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_datetime_inchour_incminute() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var dt: TDateTime; h, m, s, ms: Word;
begin
  dt := EncodeTime(10, 0, 0, 0);
  dt := IncHour(dt, 2);
  dt := IncMinute(dt, 15);
  DecodeTime(dt, h, m, s, ms);
  WriteLn(h.ToString + ':' + m.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["12:15"]);
}

#[test]
fn test_datetime_strtodate() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime; y, m, d: Word;
begin
  ShortDateFormat := 'yyyy-mm-dd';
  DateSeparator := '-';
  dt := StrToDate('2026-08-15');
  DecodeDate(dt, y, m, d);
  WriteLn(y.ToString + '-' + m.ToString + '-' + d.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["2026-8-15"]);
}

#[test]
fn test_datetime_same_date_check() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var d1, d2: TDateTime;
begin
  d1 := EncodeDateTime(2026, 5, 10, 8, 30, 0, 0);
  d2 := EncodeDateTime(2026, 5, 10, 19, 45, 0, 0);
  WriteLn(SameDate(d1, d2));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_datetime_same_time_check() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var t1, t2: TDateTime;
begin
  t1 := EncodeDateTime(2026, 1, 1, 10, 30, 0, 0);
  t2 := EncodeDateTime(2026, 6, 15, 10, 30, 0, 0);
  WriteLn(SameTime(t1, t2));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_datetime_monthsbetween() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var d1, d2: TDateTime;
begin
  d1 := EncodeDate(2026, 1, 1);
  d2 := EncodeDate(2026, 7, 1);
  WriteLn(MonthsBetween(d2, d1));
end.
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_datetime_yearsbetween() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var d1, d2: TDateTime;
begin
  d1 := EncodeDate(2020, 1, 1);
  d2 := EncodeDate(2026, 1, 1);
  WriteLn(YearsBetween(d2, d1));
end.
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_datetime_formatdatetime_custom_tokens() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var dt: TDateTime;
begin
  dt := EncodeDate(2026, 4, 9);
  WriteLn(FormatDateTime('dd/mm/yyyy', dt));
end.
"#,
    );
    assert_eq!(out, vec!["09/04/2026"]);
}

#[test]
fn test_datetime_startoftheyear() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var dt: TDateTime; y, m, d: Word;
begin
  dt := StartOfTheYear(2026);
  DecodeDate(dt, y, m, d);
  WriteLn(y.ToString + '-' + m.ToString + '-' + d.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["2026-1-1"]);
}

#[test]
fn test_datetime_endoftheyear() {
    let out = run_pascal(
        r#"
program Test;
uses DateUtils, SysUtils;
var dt: TDateTime; y, m, d: Word;
begin
  dt := EndOfTheYear(2026);
  DecodeDate(dt, y, m, d);
  WriteLn(y.ToString + '-' + m.ToString + '-' + d.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["2026-12-31"]);
}
