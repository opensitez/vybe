/// Day/month/year calendar arithmetic and date parts.
use super::helpers::run_pascal;

#[test]
fn encode_date_basic() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,2); WriteLn(FormatDateTime("yyyy-mm-dd",d)); end."#
        ),
        &["2020-01-02"]
    );
}

#[test]
fn decode_date_parts() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; y,m,day:Word; begin d:=EncodeDate(1999,12,31); DecodeDate(d,y,m,day); WriteLn(y); WriteLn(m); WriteLn(day); end."#
        ),
        &["1999", "12", "31"]
    );
}

#[test]
fn encode_time_basic() {
    assert_eq!(
        run_pascal(
            r#"program T; var t:TDateTime; begin t:=EncodeTime(1,2,3,0); WriteLn(FormatDateTime("hh:nn:ss",t)); end."#
        ),
        &["01:02:03"]
    );
}

#[test]
fn decode_time_parts() {
    assert_eq!(
        run_pascal(
            r#"program T; var t:TDateTime; h,m,s,ms:Word; begin t:=EncodeTime(23,59,58,0); DecodeTime(t,h,m,s,ms); WriteLn(h); WriteLn(m); WriteLn(s); end."#
        ),
        &["23", "59", "58"]
    );
}

#[test]
fn day_of_week_saturday() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2000,1,1); WriteLn(DayOfWeek(d)); end."#
        ),
        &["6"]
    );
}

#[test]
fn day_of_month_mid() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,6,15); WriteLn(DayOf(d)); end."#
        ),
        &["15"]
    );
}

#[test]
fn month_of_date() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,6,15); WriteLn(MonthOf(d)); end."#
        ),
        &["6"]
    );
}

#[test]
fn year_of_date() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,6,15); WriteLn(YearOf(d)); end."#
        ),
        &["2020"]
    );
}

#[test]
fn inc_month_forward() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,31); d:=IncMonth(d,1); WriteLn(FormatDateTime("yyyy-mm",d)); end."#
        ),
        &["2020-02"]
    );
}

#[test]
fn inc_month_backward() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,3,15); d:=IncMonth(d,-1); WriteLn(FormatDateTime("yyyy-mm",d)); end."#
        ),
        &["2020-02"]
    );
}

#[test]
fn inc_day_forward() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,1); d:=IncDay(d,1); WriteLn(DayOf(d)); end."#
        ),
        &["2"]
    );
}

#[test]
fn days_between_dates() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2020,1,1); b:=EncodeDate(2020,1,11); WriteLn(DaysBetween(b,a)); end."#
        ),
        &["10"]
    );
}

#[test]
fn leap_year_via_feb29() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,2,29); WriteLn(DayOf(d)=29); end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn non_leap_feb_span() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2019,2,1); b:=IncMonth(a,1); WriteLn(DaysBetween(b,a)); end."#
        ),
        &["28"]
    );
}

#[test]
fn days_in_month_jan() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2020,1,1); b:=IncMonth(a,1); WriteLn(DaysBetween(b,a)); end."#
        ),
        &["31"]
    );
}

#[test]
fn days_in_month_feb_leap() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2020,2,1); b:=IncMonth(a,1); WriteLn(DaysBetween(b,a)); end."#
        ),
        &["29"]
    );
}

#[test]
fn days_in_month_apr() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2020,4,1); b:=IncMonth(a,1); WriteLn(DaysBetween(b,a)); end."#
        ),
        &["30"]
    );
}

#[test]
fn start_of_month() {
    assert_eq!(
        run_pascal(
            r#"program T; var d,s:TDateTime; begin d:=EncodeDate(2020,6,15); s:=EncodeDate(YearOf(d),MonthOf(d),1); WriteLn(DayOf(s)); end."#
        ),
        &["1"]
    );
}

#[test]
fn end_of_month_june() {
    assert_eq!(
        run_pascal(
            r#"program T; var d,e:TDateTime; begin d:=EncodeDate(2020,6,15); e:=IncDay(IncMonth(EncodeDate(YearOf(d),MonthOf(d),1),1),-1); WriteLn(DayOf(e)); end."#
        ),
        &["30"]
    );
}

#[test]
fn date_to_str_short() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2021,3,4); WriteLn(DateToStr(d)); end."#
        ),
        &["3/4/2021"]
    );
}

#[test]
fn time_to_str() {
    assert_eq!(
        run_pascal(
            r#"program T; var t:TDateTime; begin t:=EncodeTime(14,30,0,0); WriteLn(TimeToStr(t)); end."#
        ),
        &["2:30:00 PM"]
    );
}

#[test]
fn compare_dates_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2020,5,5); b:=EncodeDate(2020,5,5); WriteLn(CompareDate(a,b)); end."#
        ),
        &["0"]
    );
}

#[test]
fn compare_dates_before() {
    assert_eq!(
        run_pascal(
            r#"program T; var early,late:TDateTime; begin early:=EncodeDate(2020,1,1); late:=EncodeDate(2020,12,31); WriteLn(CompareDate(early,late)); end."#
        ),
        &["-1"]
    );
}

#[test]
fn inc_year_style() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2019,12,31); d:=IncMonth(d,12); WriteLn(YearOf(d)); end."#
        ),
        &["2020"]
    );
}

#[test]
fn day_of_week_monday() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2021,3,1); WriteLn(DayOfWeek(d)); end."#
        ),
        &["1"]
    );
}

#[test]
fn hour_of_time() {
    assert_eq!(
        run_pascal(
            r#"program T; var t:TDateTime; begin t:=EncodeTime(15,0,0,0); WriteLn(HourOf(t)); end."#
        ),
        &["15"]
    );
}

#[test]
fn minute_of_time() {
    assert_eq!(
        run_pascal(
            r#"program T; var t:TDateTime; begin t:=EncodeTime(15,45,0,0); WriteLn(MinuteOf(t)); end."#
        ),
        &["45"]
    );
}

#[test]
fn second_of_time() {
    assert_eq!(
        run_pascal(
            r#"program T; var t:TDateTime; begin t:=EncodeTime(0,0,33,0); WriteLn(SecondOf(t)); end."#
        ),
        &["33"]
    );
}

#[test]
fn encode_date_leap_day() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,2,29); WriteLn(FormatDateTime("mm-dd",d)); end."#
        ),
        &["02-29"]
    );
}

#[test]
fn add_days_week() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,1); d:=IncDay(d,7); WriteLn(FormatDateTime("yyyy-mm-dd",d)); end."#
        ),
        &["2020-01-08"]
    );
}

#[test]
fn subtract_one_day() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,2); d:=IncDay(d,-1); WriteLn(DayOf(d)); end."#
        ),
        &["1"]
    );
}

#[test]
fn inc_month_two() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,15); d:=IncMonth(d,2); WriteLn(MonthOf(d)); end."#
        ),
        &["3"]
    );
}

#[test]
fn inc_month_neg_three() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,6,15); d:=IncMonth(d,-3); WriteLn(MonthOf(d)); end."#
        ),
        &["3"]
    );
}

#[test]
fn day_of_year_jan1() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,1,1); WriteLn(DaysBetween(EncodeDate(2020,1,1),d)+1); end."#
        ),
        &["1"]
    );
}

#[test]
fn day_of_year_dec31() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,12,31); WriteLn(DaysBetween(EncodeDate(2020,1,1),d)+1); end."#
        ),
        &["366"]
    );
}

#[test]
fn weeks_between() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2020,1,1); b:=EncodeDate(2020,1,15); WriteLn(DaysBetween(b,a) div 7); end."#
        ),
        &["2"]
    );
}

#[test]
fn same_month_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,7,4); if MonthOf(d)=7 then WriteLn("july"); end."#
        ),
        &["july"]
    );
}

#[test]
fn quarter_from_month() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2020,11,1); WriteLn((MonthOf(d)-1) div 3 + 1); end."#
        ),
        &["4"]
    );
}

#[test]
fn date_plus_time() {
    assert_eq!(
        run_pascal(
            r#"program T; var d,t,dt:TDateTime; begin d:=EncodeDate(2020,1,1); t:=EncodeTime(12,0,0,0); dt:=d+t; WriteLn(FormatDateTime("yyyy-mm-dd",dt)); WriteLn(FormatDateTime("hh:nn",dt)); end."#
        ),
        &["2020-01-01", "12:00"]
    );
}

#[test]
fn same_date_check() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2001,6,15); b:=EncodeDate(2001,6,15); WriteLn(SameDate(a,b)); end."#
        ),
        &["true"]
    );
}

#[test]
fn calendar_age_days() {
    assert_eq!(
        run_pascal(
            r#"program T; var birth,today:TDateTime; begin birth:=EncodeDate(2000,1,1); today:=EncodeDate(2000,1,31); WriteLn(DaysBetween(today,birth)); end."#
        ),
        &["30"]
    );
}

#[test]
fn last_day_feb_nonleap() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:TDateTime; begin a:=EncodeDate(2019,2,1); b:=IncMonth(a,1); WriteLn(DaysBetween(b,a)); end."#
        ),
        &["28"]
    );
}
