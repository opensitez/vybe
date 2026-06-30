/// Date/Time/Now/EncodeDate/DecodeDate and related — beyond test_time.rs.
use super::helpers::run_pascal;

#[test]
fn encode_date_jan_first_2000() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,1,1); WriteLn(FormatDateTime('yyyy-mm-dd', d)); end."#),
        &["2000-01-01"]
    );
}

#[test]
fn decode_date_roundtrip() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; y,m,day:Word; begin d:=EncodeDate(2012,7,4); DecodeDate(d,y,m,day); WriteLn(y); WriteLn(m); WriteLn(day); end."#),
        &["2012", "7", "4"]
    );
}

#[test]
fn encode_time_midnight() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=EncodeTime(0,0,0,0); WriteLn(FormatDateTime('hh:nn:ss', t)); end."#),
        &["00:00:00"]
    );
}

#[test]
fn decode_time_components() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; h,m,s,ms:Word; begin t:=EncodeTime(15,45,30,0); DecodeTime(t,h,m,s,ms); WriteLn(h); WriteLn(m); WriteLn(s); end."#),
        &["15", "45", "30"]
    );
}

#[test]
fn date_add_one_day() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,1,1); d:=d+1; WriteLn(FormatDateTime('yyyy-mm-dd', d)); end."#),
        &["2000-01-02"]
    );
}

#[test]
fn date_subtract_one_day() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,1,2); d:=d-1; WriteLn(FormatDateTime('yyyy-mm-dd', d)); end."#),
        &["2000-01-01"]
    );
}

#[test]
fn days_between_two_dates() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2000,1,1); b:=EncodeDate(2000,1,11); WriteLn(Trunc(b-a)); end."#),
        &["10"]
    );
}

#[test]
fn compare_date_less_than() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(1999,12,31); b:=EncodeDate(2000,1,1); WriteLn(CompareDate(a,b)); end."#),
        &["-1"]
    );
}

#[test]
fn compare_date_equal() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2010,5,5); b:=EncodeDate(2010,5,5); WriteLn(CompareDate(a,b)); end."#),
        &["0"]
    );
}

#[test]
fn same_date_true() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2001,6,15); b:=EncodeDate(2001,6,15); WriteLn(SameDate(a,b)); end."#),
        &["true"]
    );
}

#[test]
fn same_date_false_adjacent() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2001,6,15); b:=EncodeDate(2001,6,16); WriteLn(SameDate(a,b)); end."#),
        &["false"]
    );
}

#[test]
fn day_of_week_saturday() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,1,1); WriteLn(DayOfWeek(d)); end."#),
        &["6"]
    );
}

#[test]
fn day_of_month_last_jan() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,31); WriteLn(DayOf(d)); end."#),
        &["31"]
    );
}

#[test]
fn month_of_july() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2015,7,4); WriteLn(MonthOf(d)); end."#),
        &["7"]
    );
}

#[test]
fn year_of_date() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2015,7,4); WriteLn(YearOf(d)); end."#),
        &["2015"]
    );
}

#[test]
fn leap_day_feb_29_2000() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,2,29); WriteLn(FormatDateTime('yyyy-mm-dd', d)); end."#),
        &["2000-02-29"]
    );
}

#[test]
fn date_to_str_format() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(1999,1,2); WriteLn(DateToStr(d)); end."#),
        &["1/2/1999"]
    );
}

#[test]
fn time_to_str_noon() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=EncodeTime(12,0,0,0); WriteLn(TimeToStr(t)); end."#),
        &["12:00:00 PM"]
    );
}

#[test]
fn hour_of_afternoon() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=EncodeTime(14,0,0,0); WriteLn(HourOf(t)); end."#),
        &["14"]
    );
}

#[test]
fn minute_of_time() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=EncodeTime(9,27,0,0); WriteLn(MinuteOf(t)); end."#),
        &["27"]
    );
}

#[test]
fn second_of_time() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=EncodeTime(9,27,53,0); WriteLn(SecondOf(t)); end."#),
        &["53"]
    );
}

#[test]
fn encode_date_time_combined_compare() {
    assert_eq!(
        run_pascal(r#"program T; var d,t,dt:TDateTime; begin d:=EncodeDate(2000,1,1); t:=EncodeTime(1,2,3,0); dt:=d+t; WriteLn(FormatDateTime('yyyy-mm-dd', dt)); WriteLn(FormatDateTime('hh:nn:ss', dt)); end."#),
        &["2000-01-01", "01:02:03"]
    );
}

#[test]
fn inc_month_add_one() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,1,31); d:=IncMonth(d,1); WriteLn(FormatDateTime('yyyy-mm', d)); end."#),
        &["2000-02"]
    );
}

#[test]
fn inc_month_subtract_one() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,3,15); d:=IncMonth(d,-1); WriteLn(FormatDateTime('yyyy-mm', d)); end."#),
        &["2000-02"]
    );
}

#[test]
fn first_of_month_via_encode_date() {
    assert_eq!(
        run_pascal(r#"program T; var d,s:TDateTime; begin d:=EncodeDate(2015,6,15); s:=EncodeDate(YearOf(d),MonthOf(d),1); WriteLn(DayOf(s)); end."#),
        &["1"]
    );
}

#[test]
fn last_of_month_via_inc_day() {
    assert_eq!(
        run_pascal(r#"program T; var d,e:TDateTime; begin d:=EncodeDate(2020,2,1); e:=IncDay(IncMonth(d,1),-1); WriteLn(DayOf(e)); end."#),
        &["29"]
    );
}

#[test]
fn compare_time_earlier() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeTime(8,0,0,0); b:=EncodeTime(9,0,0,0); WriteLn(CompareTime(a,b)); end."#),
        &["-1"]
    );
}

#[test]
fn leap_year_detected_via_feb_twenty_nine() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,2,29); WriteLn(DayOf(d)=29); end."#),
        &["true"]
    );
}

#[test]
fn non_leap_year_feb_has_twenty_eight_days() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(1900,2,1); b:=IncMonth(a,1); WriteLn(DaysBetween(b,a)); end."#),
        &["28"]
    );
}

#[test]
fn days_in_february_leap_via_month_span() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2000,2,1); b:=IncMonth(a,1); WriteLn(DaysBetween(b,a)); end."#),
        &["29"]
    );
}

#[test]
fn days_in_year_leap_via_jan_first_span() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2000,1,1); b:=EncodeDate(2001,1,1); WriteLn(DaysBetween(b,a)); end."#),
        &["366"]
    );
}

#[test]
fn str_to_date_parses() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=StrToDate('3/15/2010'); WriteLn(YearOf(d)); WriteLn(MonthOf(d)); WriteLn(DayOf(d)); end."#),
        &["2010", "3", "15"]
    );
}

#[test]
fn str_to_time_parses_hour() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=StrToTime('08:30:00'); WriteLn(HourOf(t)); end."#),
        &["8"]
    );
}

#[test]
fn inc_hour_advances_clock() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=EncodeTime(10,0,0,0); t:=IncHour(t,2); WriteLn(HourOf(t)); end."#),
        &["12"]
    );
}

#[test]
fn inc_minute_advances_clock() {
    assert_eq!(
        run_pascal(r#"program T; var t:TDateTime; begin t:=EncodeTime(10,15,0,0); t:=IncMinute(t,30); WriteLn(MinuteOf(t)); end."#),
        &["45"]
    );
}

#[test]
fn date_plus_time_preserves_date_part() {
    assert_eq!(
        run_pascal(r#"program T; var d,t,dt:TDateTime; begin d:=EncodeDate(2010,1,2); t:=EncodeTime(15,0,0,0); dt:=d+t; WriteLn(YearOf(dt)); WriteLn(MonthOf(dt)); WriteLn(DayOf(dt)); end."#),
        &["2010", "1", "2"]
    );
}

#[test]
fn date_plus_time_preserves_time_part() {
    assert_eq!(
        run_pascal(r#"program T; var d,t,dt:TDateTime; begin d:=EncodeDate(2010,1,2); t:=EncodeTime(15,30,45,0); dt:=d+t; WriteLn(HourOf(dt)); WriteLn(MinuteOf(dt)); WriteLn(SecondOf(dt)); end."#),
        &["15", "30", "45"]
    );
}

#[test]
fn minutes_between_two_times() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:TDateTime; begin a:=EncodeTime(10,0,0,0); b:=EncodeTime(10,45,0,0); WriteLn(Trunc(MinutesBetween(b,a))); end."#),
        &["45"]
    );
}

#[test]
fn encode_decode_year_boundary() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; y,m,day:Word; begin d:=EncodeDate(1999,12,31); d:=d+1; DecodeDate(d,y,m,day); WriteLn(y); WriteLn(m); WriteLn(day); end."#),
        &["2000", "1", "1"]
    );
}

#[test]
fn day_of_week_sunday_on_known_date() {
    assert_eq!(
        run_pascal(r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,1,2); WriteLn(DayOfWeek(d)); end."#),
        &["7"]
    );
}

#[test]
fn day_of_year_last_day_via_days_between() {
    assert_eq!(
        run_pascal(r#"program T; var start,last:TDateTime; begin start:=EncodeDate(2000,1,1); last:=EncodeDate(2000,12,31); WriteLn(DaysBetween(last,start)+1); end."#),
        &["366"]
    );
}
