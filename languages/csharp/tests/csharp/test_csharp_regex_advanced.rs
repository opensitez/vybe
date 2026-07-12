//! Advanced `Regex`: named groups, Matches (all), Replace with evaluator, lookahead.
use super::helpers::run_csharp;

#[test]
fn named_group_captured_by_name() {
    assert_eq!(
        run_csharp(
            r#"var m = System.Text.RegularExpressions.Regex.Match("date=2024-06-15", @"(?<year>\d{4})-(?<month>\d{2})");
Console.WriteLine(m.Groups["year"].Value);
Console.WriteLine(m.Groups["month"].Value);"#
        ),
        &["2024", "06"]
    );
}

#[test]
fn matches_returns_all_non_overlapping_occurrences() {
    assert_eq!(
        run_csharp(
            r#"var matches = System.Text.RegularExpressions.Regex.Matches("a1 b2 c3", @"\d");
Console.WriteLine(matches.Count);"#
        ),
        &["3"]
    );
}

#[test]
fn replace_with_match_evaluator_transforms_each_match() {
    assert_eq!(
        run_csharp(
            r#"string result = System.Text.RegularExpressions.Regex.Replace(
    "a1b2c3", @"\d",
    m => ((int.Parse(m.Value)*2)).ToString());
Console.WriteLine(result);"#
        ),
        &["a2b4c6"]
    );
}

#[test]
fn anchored_pattern_does_not_match_mid_string() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(System.Text.RegularExpressions.Regex.IsMatch("abc", @"^\d+$"));"#
        ),
        &["False"]
    );
}

#[test]
fn character_class_matches_any_listed_char() {
    assert_eq!(
        run_csharp(
            r#"var m = System.Text.RegularExpressions.Regex.Match("hello", @"[aeiou]");
Console.WriteLine(m.Value);"#
        ),
        &["e"]
    );
}

#[test]
fn quantifier_plus_requires_one_or_more_digits() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(System.Text.RegularExpressions.Regex.IsMatch("007", @"^\d+$"));"#
        ),
        &["True"]
    );
}

#[test]
fn multiline_option_applies_caret_to_each_line_start() {
    assert_eq!(
        run_csharp(
            r#"var matches = System.Text.RegularExpressions.Regex.Matches(
    "start\nnew line", @"^[a-z]",
    System.Text.RegularExpressions.RegexOptions.Multiline);
Console.WriteLine(matches.Count);"#
        ),
        &["2"]
    );
}
