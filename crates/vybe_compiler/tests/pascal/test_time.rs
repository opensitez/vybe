//! Date and time: EncodeDate, DecodeDate, calendar arithmetic.
use super::helpers::run_pascal;

#[test]
fn encode_date_year_month_day() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2020, 1, 2);
  WriteLn(FormatDateTime('yyyy-mm-dd', d));
end."#
        ),
        &["2020-01-02"]
    );
}

#[test]
fn decode_date_splits_encoded_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime; y, m, day: Word;
begin
  d := EncodeDate(1999, 12, 31);
  DecodeDate(d, y, m, day);
  WriteLn(y);
  WriteLn(m);
  WriteLn(day);
end."#
        ),
        &["1999", "12", "31"]
    );
}

#[test]
fn encode_time_hour_minute_second() {
    assert_eq!(
        run_pascal(
            r#"program T;
var t: TDateTime;
begin
  t := EncodeTime(1, 2, 3, 0);
  WriteLn(FormatDateTime('hh:nn:ss', t));
end."#
        ),
        &["01:02:03"]
    );
}

#[test]
fn decode_time_reads_components() {
    assert_eq!(
        run_pascal(
            r#"program T;
var t: TDateTime; h, m, s, ms: Word;
begin
  t := EncodeTime(23, 59, 58, 0);
  DecodeTime(t, h, m, s, ms);
  WriteLn(h);
  WriteLn(m);
  WriteLn(s);
end."#
        ),
        &["23", "59", "58"]
    );
}

#[test]
fn day_of_week_for_known_date() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2000, 1, 1);
  WriteLn(DayOfWeek(d));
end."#
        ),
        &["6"]
    );
}

#[test]
fn days_between_two_dates() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: TDateTime;
begin
  a := EncodeDate(2000, 1, 1);
  b := EncodeDate(2000, 1, 11);
  WriteLn(DaysBetween(b, a));
end."#
        ),
        &["10"]
    );
}

#[test]
fn inc_day_advances_calendar() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2000, 1, 31);
  d := IncDay(d, 1);
  WriteLn(FormatDateTime('yyyy-mm-dd', d));
end."#
        ),
        &["2000-02-01"]
    );
}

#[test]
fn compare_date_equal_returns_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: TDateTime;
begin
  a := EncodeDate(2010, 5, 5);
  b := EncodeDate(2010, 5, 5);
  WriteLn(CompareDate(a, b));
end."#
        ),
        &["0"]
    );
}

#[test]
fn year_of_extracts_calendar_year() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2015, 7, 4);
  WriteLn(YearOf(d));
end."#
        ),
        &["2015"]
    );
}

#[test]
fn month_of_extracts_calendar_month() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2015, 7, 4);
  WriteLn(MonthOf(d));
end."#
        ),
        &["7"]
    );
}

#[test]
fn day_of_extracts_calendar_day() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2015, 7, 4);
  WriteLn(DayOf(d));
end."#
        ),
        &["4"]
    );
}

#[test]
fn inc_month_crosses_year_boundary() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2015, 12, 15);
  d := IncMonth(d, 2);
  WriteLn(FormatDateTime('yyyy-mm', d));
end."#
        ),
        &["2016-02"]
    );
}

#[test]
fn dec_month_crosses_year_backward() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2015, 1, 10);
  d := IncMonth(d, -1);
  WriteLn(FormatDateTime('yyyy-mm', d));
end."#
        ),
        &["2014-12"]
    );
}

#[test]
fn same_date_true_for_identical_dates() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: TDateTime;
begin
  a := EncodeDate(2001, 6, 15);
  b := EncodeDate(2001, 6, 15);
  WriteLn(SameDate(a, b));
end."#
        ),
        &["true"]
    );
}

#[test]
fn same_date_false_when_days_differ() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: TDateTime;
begin
  a := EncodeDate(2001, 6, 15);
  b := EncodeDate(2001, 6, 16);
  WriteLn(SameDate(a, b));
end."#
        ),
        &["false"]
    );
}

#[test]
fn compare_date_orders_chronologically() {
    assert_eq!(
        run_pascal(
            r#"program T;
var early, late: TDateTime;
begin
  early := EncodeDate(2000, 1, 1);
  late := EncodeDate(2001, 1, 1);
  WriteLn(CompareDate(early, late));
end."#
        ),
        &["-1"]
    );
}

#[test]
fn encode_date_leap_year_feb_twenty_nine() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2000, 2, 29);
  WriteLn(FormatDateTime('yyyy-mm-dd', d));
end."#
        ),
        &["2000-02-29"]
    );
}

#[test]
fn hours_between_two_times_on_same_day() {
    assert_eq!(
        run_pascal(
            r#"program T;
var t1, t2: TDateTime;
begin
  t1 := EncodeTime(10, 0, 0, 0);
  t2 := EncodeTime(13, 30, 0, 0);
  WriteLn(Trunc(HoursBetween(t2, t1)));
end."#
        ),
        &["3"]
    );
}

#[test]
fn date_to_str_uses_short_format() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(1999, 1, 2);
  WriteLn(DateToStr(d));
end."#
        ),
        &["1/2/1999"]
    );
}

#[test]
fn str_to_date_parses_slash_form() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := StrToDate('3/15/2010');
  WriteLn(YearOf(d));
  WriteLn(MonthOf(d));
  WriteLn(DayOf(d));
end."#
        ),
        &["2010", "3", "15"]
    );
}

#[test]
fn encode_time_midnight_components() {
    assert_eq!(
        run_pascal(
            r#"program T;
var t: TDateTime;
begin
  t := EncodeTime(0, 0, 0, 0);
  WriteLn(HourOf(t));
  WriteLn(MinuteOf(t));
end."#
        ),
        &["0", "0"]
    );
}

#[test]
fn decode_time_splits_hours_minutes() {
    assert_eq!(
        run_pascal(
            r#"program T;
var t: TDateTime;
    h, m, s, ms: Word;
begin
  t := EncodeTime(14, 25, 30, 0);
  DecodeTime(t, h, m, s, ms);
  WriteLn(h);
  WriteLn(m);
end."#
        ),
        &["14", "25"]
    );
}

#[test]
fn inc_day_advances_calendar_date() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2020, 1, 31);
  d := IncDay(d, 1);
  WriteLn(DayOf(d));
end."#
        ),
        &["1"]
    );
}

#[test]
fn days_between_same_date_is_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: TDateTime;
begin
  d := EncodeDate(2021, 6, 1);
  WriteLn(DaysBetween(d, d));
end."#
        ),
        &["0"]
    );
}

#[test]
fn compare_date_orders_earlier_first() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: TDateTime;
begin
  a := EncodeDate(2000, 1, 1);
  b := EncodeDate(2000, 1, 2);
  WriteLn(CompareDate(a, b));
end."#
        ),
        &["-1"]
    );
}


