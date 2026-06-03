use super::helpers::run_vb;

macro_rules! vb_expr_spec {
    ($name:ident, $expr:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let src = format!(
                r#"Module M
    Sub Main()
        Console.WriteLine({})
    End Sub
End Module
"#,
                $expr
            );
            let out = run_vb(&src);
            assert_eq!(out, vec![super::helpers::dotnet_expected_one($expected)]);
        }
    };
}

vb_expr_spec!(
    datetime_spec_dateserial_builds_first_day_of_year,
    r#"CStr(DateSerial(2024, 1, 1))"#,
    "1/1/2024"
);
vb_expr_spec!(
    datetime_spec_dateserial_builds_leap_day,
    r#"CStr(DateSerial(2024, 2, 29))"#,
    "2/29/2024"
);
vb_expr_spec!(
    datetime_spec_dateserial_rolls_over_day_zero_to_previous_month,
    r#"CStr(DateSerial(2024, 3, 0))"#,
    "2/29/2024"
);
vb_expr_spec!(
    datetime_spec_timeserial_builds_morning_time,
    r#"CStr(TimeSerial(9, 30, 0))"#,
    "9:30:00 AM"
);
vb_expr_spec!(
    datetime_spec_timeserial_builds_evening_time,
    r#"CStr(TimeSerial(20, 15, 0))"#,
    "8:15:00 PM"
);
vb_expr_spec!(
    datetime_spec_timeserial_rolls_minutes_into_next_hour,
    r#"CStr(TimeSerial(10, 75, 0))"#,
    "11:15:00 AM"
);
vb_expr_spec!(
    datetime_spec_dateadd_adds_days_to_literal_date,
    r#"CStr(DateAdd(DateInterval.Day, 5, #5/14/2024#))"#,
    "5/19/2024"
);
vb_expr_spec!(
    datetime_spec_dateadd_adds_months_to_literal_date,
    r#"CStr(DateAdd(DateInterval.Month, 2, #5/14/2024#))"#,
    "7/14/2024"
);
vb_expr_spec!(
    datetime_spec_dateadd_adds_years_to_literal_date,
    r#"CStr(DateAdd(DateInterval.Year, 1, #5/14/2024#))"#,
    "5/14/2025"
);
vb_expr_spec!(
    datetime_spec_dateadd_adds_hours_to_datetime_value,
    r#"CStr(DateAdd(DateInterval.Hour, 3, #5/14/2024 1:00 PM#))"#,
    "5/14/2024 4:00:00 PM"
);
vb_expr_spec!(
    datetime_spec_dateadd_subtracts_days_with_negative_amount,
    r#"CStr(DateAdd(DateInterval.Day, -7, #5/14/2024#))"#,
    "5/7/2024"
);
vb_expr_spec!(
    datetime_spec_datediff_returns_day_span_between_dates,
    r#"DateDiff(DateInterval.Day, #5/1/2024#, #5/14/2024#)"#,
    "13"
);
vb_expr_spec!(
    datetime_spec_datediff_returns_month_span_between_dates,
    r#"DateDiff(DateInterval.Month, #1/1/2024#, #5/14/2024#)"#,
    "4"
);
vb_expr_spec!(
    datetime_spec_datediff_returns_year_span_between_dates,
    r#"DateDiff(DateInterval.Year, #1/1/2020#, #5/14/2024#)"#,
    "4"
);
vb_expr_spec!(
    datetime_spec_datediff_returns_hour_span_between_times,
    r#"DateDiff(DateInterval.Hour, #5/14/2024 1:00 AM#, #5/14/2024 8:00 AM#)"#,
    "7"
);
vb_expr_spec!(
    datetime_spec_year_extracts_year_component,
    r#"Year(#5/14/2024#)"#,
    "2024"
);
vb_expr_spec!(
    datetime_spec_month_extracts_month_component,
    r#"Month(#5/14/2024#)"#,
    "5"
);
vb_expr_spec!(
    datetime_spec_day_extracts_day_component,
    r#"Day(#5/14/2024#)"#,
    "14"
);
vb_expr_spec!(
    datetime_spec_hour_extracts_hour_component,
    r#"Hour(#5/14/2024 3:45 PM#)"#,
    "15"
);
vb_expr_spec!(
    datetime_spec_minute_extracts_minute_component,
    r#"Minute(#5/14/2024 3:45 PM#)"#,
    "45"
);
vb_expr_spec!(
    datetime_spec_second_extracts_second_component,
    r#"Second(#5/14/2024 3:45:59 PM#)"#,
    "59"
);
vb_expr_spec!(
    datetime_spec_weekday_returns_numeric_day_index,
    r#"Weekday(#5/13/2024#)"#,
    "2"
);
vb_expr_spec!(
    datetime_spec_weekdayname_returns_full_day_name,
    r#"WeekdayName(2)"#,
    "Monday"
);
vb_expr_spec!(
    datetime_spec_weekdayname_returns_abbreviated_day_name,
    r#"WeekdayName(2, True)"#,
    "Mon"
);
vb_expr_spec!(
    datetime_spec_monthname_returns_full_month_name,
    r#"MonthName(1)"#,
    "January"
);
vb_expr_spec!(
    datetime_spec_monthname_returns_abbreviated_month_name,
    r#"MonthName(1, True)"#,
    "Jan"
);
vb_expr_spec!(
    datetime_spec_datevalue_parses_short_date_string,
    r#"CStr(DateValue("5/14/2024"))"#,
    "5/14/2024"
);
vb_expr_spec!(
    datetime_spec_timevalue_parses_short_time_string,
    r#"CStr(TimeValue("3:45 PM"))"#,
    "3:45:00 PM"
);
vb_expr_spec!(
    datetime_spec_cdate_parses_datetime_string,
    r#"CStr(CDate("5/14/2024 3:45:59 PM"))"#,
    "5/14/2024 3:45:59 PM"
);
vb_expr_spec!(
    datetime_spec_isdate_accepts_valid_date_string,
    r#"IsDate("5/14/2024")"#,
    "true"
);
vb_expr_spec!(
    datetime_spec_isdate_rejects_invalid_date_string,
    r#"IsDate("not-a-date")"#,
    "false"
);
vb_expr_spec!(
    datetime_spec_now_returns_non_nothing_value,
    r#"Not IsNothing(Now)"#,
    "true"
);
vb_expr_spec!(
    datetime_spec_today_returns_non_nothing_value,
    r#"Not IsNothing(Today)"#,
    "true"
);
vb_expr_spec!(
    datetime_spec_timeofday_returns_non_nothing_value,
    r#"Not IsNothing(TimeOfDay)"#,
    "true"
);
vb_expr_spec!(
    datetime_spec_clng_converts_integer_text,
    r#"CLng("42")"#,
    "42"
);
vb_expr_spec!(
    datetime_spec_cshort_converts_small_integer_text,
    r#"CShort("12")"#,
    "12"
);
vb_expr_spec!(
    datetime_spec_cbyte_converts_small_positive_integer,
    r#"CByte("255")"#,
    "255"
);
vb_expr_spec!(
    datetime_spec_csng_converts_integer_literal_to_single,
    r#"CSng(5)"#,
    "5"
);
vb_expr_spec!(
    datetime_spec_cdec_converts_decimal_text,
    r#"CDec("12.75")"#,
    "12.75"
);
vb_expr_spec!(
    datetime_spec_cchar_converts_numeric_code_to_character,
    r#"CChar(ChrW(65))"#,
    "A"
);
vb_expr_spec!(
    datetime_spec_fix_truncates_negative_fraction_toward_zero,
    r#"Fix(-3.8)"#,
    "-3"
);
vb_expr_spec!(
    datetime_spec_int_rounds_negative_fraction_downward,
    r#"Int(-3.8)"#,
    "-4"
);
vb_expr_spec!(
    datetime_spec_hex_formats_integer_as_hex_text,
    r#"Hex(255)"#,
    "FF"
);
vb_expr_spec!(
    datetime_spec_oct_formats_integer_as_octal_text,
    r#"Oct(64)"#,
    "100"
);
vb_expr_spec!(
    datetime_spec_sgn_returns_negative_indicator,
    r#"Sgn(-10)"#,
    "-1"
);
vb_expr_spec!(datetime_spec_sgn_returns_zero_indicator, r#"Sgn(0)"#, "0");
vb_expr_spec!(
    datetime_spec_sgn_returns_positive_indicator,
    r#"Sgn(10)"#,
    "1"
);
vb_expr_spec!(
    datetime_spec_dateserial_and_year_roundtrip_same_year,
    r#"Year(DateSerial(2030, 6, 15))"#,
    "2030"
);
vb_expr_spec!(
    datetime_spec_timeserial_and_hour_roundtrip_same_hour,
    r#"Hour(TimeSerial(22, 5, 0))"#,
    "22"
);
vb_expr_spec!(
    datetime_spec_dateadd_can_chain_multiple_intervals,
    r#"CStr(DateAdd(DateInterval.Hour, 2, DateAdd(DateInterval.Day, 1, #5/14/2024 10:00 AM#)))"#,
    "5/15/2024 12:00:00 PM"
);
vb_expr_spec!(
    datetime_spec_datepart_returns_quarter_number,
    r#"DatePart(DateInterval.Quarter, #5/14/2024#)"#,
    "2"
);
vb_expr_spec!(
    datetime_spec_datepart_returns_day_of_year,
    r#"DatePart(DateInterval.DayOfYear, #12/31/2024#)"#,
    "366"
);
