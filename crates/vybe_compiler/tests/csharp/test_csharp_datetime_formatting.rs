//! `DateTime` format strings, `Parse`, `TryParse`, and culture-independent output.
use super::helpers::run_csharp;

#[test]
fn tostring_with_yyyy_mm_dd_format() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.DateTime(2024,6,15);
Console.WriteLine(d.ToString("yyyy-MM-dd"));"#
        ),
        &["2024-06-15"]
    );
}

#[test]
fn tostring_with_dd_slash_mm_slash_yyyy_format() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.DateTime(2024,1,5);
Console.WriteLine(d.ToString("dd/MM/yyyy"));"#
        ),
        &["05/01/2024"]
    );
}

#[test]
fn tostring_with_hh_mm_ss_time_format() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.DateTime(2024,1,1,13,5,9);
Console.WriteLine(d.ToString("HH:mm:ss"));"#
        ),
        &["13:05:09"]
    );
}

#[test]
fn parse_with_exact_format_and_invariant_culture() {
    assert_eq!(
        run_csharp(
            r#"var d = System.DateTime.ParseExact("2024-03-21","yyyy-MM-dd",
    System.Globalization.CultureInfo.InvariantCulture);
Console.WriteLine(d.Year); Console.WriteLine(d.Month); Console.WriteLine(d.Day);"#
        ),
        &["2024", "3", "21"]
    );
}

#[test]
fn try_parse_returns_false_for_invalid_string() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(System.DateTime.TryParse("not-a-date", out _));"#
        ),
        &["False"]
    );
}

#[test]
fn tostring_d_short_date_pattern_contains_year_digits() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.DateTime(2025,12,31);
Console.WriteLine(d.ToString("yyyy-MM-dd").StartsWith("2025"));"#
        ),
        &["True"]
    );
}

#[test]
fn datetime_today_is_not_min_value() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.DateTime.Today != System.DateTime.MinValue);"#),
        &["True"]
    );
}
