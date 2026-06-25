//! `System.Text.RegularExpressions.Regex` matching and capture groups.
use super::helpers::run_csharp;

#[test]
fn regex_is_match_reports_success_for_literal_pattern() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Text.RegularExpressions.Regex.IsMatch("abc123", @"\d+"));"#),
        &["True"]
    );
}

#[test]
fn regex_match_value_returns_first_captured_group() {
    assert_eq!(
        run_csharp(
            r#"
var match = System.Text.RegularExpressions.Regex.Match("id=42", @"id=(\d+)");
Console.WriteLine(match.Groups[1].Value);
"#
        ),
        &["42"]
    );
}

#[test]
fn regex_replace_substitutes_all_occurrences_with_replacement_text() {
    assert_eq!(
        run_csharp(
            r#"
var text = System.Text.RegularExpressions.Regex.Replace("a-b-c", "-", "_");
Console.WriteLine(text);
"#
        ),
        &["a_b_c"]
    );
}

#[test]
fn regex_split_returns_segments_around_delimiter_pattern() {
    assert_eq!(
        run_csharp(
            r#"
var parts = System.Text.RegularExpressions.Regex.Split("one,two,three", ",");
Console.WriteLine(parts[1]);
"#
        ),
        &["two"]
    );
}

#[test]
fn regex_options_ignore_case_matches_differing_casing() {
    assert_eq!(
        run_csharp(
            r#"
bool ok = System.Text.RegularExpressions.Regex.IsMatch(
    "Hello",
    "hello",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
Console.WriteLine(ok);
"#
        ),
        &["True"]
    );
}
