//! Basic `System.Text.RegularExpressions.Regex` operations not covered in advanced file.
use super::helpers::run_csharp;

#[test]
fn regex_is_match_returns_true_for_pattern_found() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Text.RegularExpressions.Regex.IsMatch("hello123","[0-9]+"));"#),
        &["True"]
    );
}

#[test]
fn regex_is_match_returns_false_when_no_match() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Text.RegularExpressions.Regex.IsMatch("hello","^[0-9]+$"));"#),
        &["False"]
    );
}

#[test]
fn regex_match_extracts_first_occurrence() {
    assert_eq!(
        run_csharp(r#"var m=System.Text.RegularExpressions.Regex.Match("abc123def","[0-9]+");
Console.WriteLine(m.Value);"#),
        &["123"]
    );
}

#[test]
fn regex_replace_substitutes_pattern_occurrences() {
    assert_eq!(
        run_csharp(r##"string r=System.Text.RegularExpressions.Regex.Replace("a1b2c3","[0-9]","#");
Console.WriteLine(r);"##),
        &["a#b#c#"]
    );
}

#[test]
fn regex_split_divides_on_pattern() {
    assert_eq!(
        run_csharp(r#"var parts=System.Text.RegularExpressions.Regex.Split("one1two2three","[0-9]");
Console.WriteLine(parts.Length); Console.WriteLine(parts[1]);"#),
        &["3", "two"]
    );
}

#[test]
fn regex_options_ignore_case_matches_mixed() {
    assert_eq!(
        run_csharp(r#"bool r=System.Text.RegularExpressions.Regex.IsMatch("Hello","hello",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
Console.WriteLine(r);"#),
        &["True"]
    );
}
