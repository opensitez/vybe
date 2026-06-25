//! Advanced DateTime: arithmetic, comparison, DayOfWeek, leap years, ticks.
use super::helpers::run_csharp;

#[test]
fn datetime_add_days_correctly_crosses_month_boundary() {
    assert_eq!(
        run_csharp(r#"var d=new System.DateTime(2024,1,30).AddDays(3);
Console.WriteLine(d.Month); Console.WriteLine(d.Day);"#),
        &["2", "2"]
    );
}

#[test]
fn datetime_subtract_yields_timespan() {
    assert_eq!(
        run_csharp(r#"var a=new System.DateTime(2024,1,10);
var b=new System.DateTime(2024,1,1);
var diff=a-b;
Console.WriteLine(diff.Days);"#),
        &["9"]
    );
}

#[test]
fn datetime_day_of_week_is_correct() {
    assert_eq!(
        run_csharp(r#"var d=new System.DateTime(2024,1,1);
Console.WriteLine(d.DayOfWeek);"#),
        &["Monday"]
    );
}

#[test]
fn datetime_is_leap_year_true_for_divisible_by_4() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.DateTime.IsLeapYear(2024));
Console.WriteLine(System.DateTime.IsLeapYear(2023));"#),
        &["True", "False"]
    );
}

#[test]
fn datetime_days_in_month_february_leap() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.DateTime.DaysInMonth(2024,2));
Console.WriteLine(System.DateTime.DaysInMonth(2023,2));"#),
        &["29", "28"]
    );
}

#[test]
fn timespan_total_minutes_converts_hours_and_minutes() {
    assert_eq!(
        run_csharp(r#"var ts=new System.TimeSpan(2,30,0);
Console.WriteLine(ts.TotalMinutes);"#),
        &["150"]
    );
}

#[test]
fn datetime_min_value_is_year_1() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.DateTime.MinValue.Year);"#),
        &["1"]
    );
}
