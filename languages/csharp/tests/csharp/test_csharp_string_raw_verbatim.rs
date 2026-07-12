//! Verbatim string literals, raw string literals, and escape sequences.
use super::helpers::run_csharp;

#[test]
fn verbatim_string_preserves_backslashes() {
    assert_eq!(
        run_csharp(
            r#"string path=@"C:\Users\test\file.txt";
Console.WriteLine(path.Contains(@"\"));"#
        ),
        &["True"]
    );
}

#[test]
fn verbatim_string_preserves_newlines_in_literal() {
    assert_eq!(
        run_csharp("string s=@\"line1\nline2\";\nConsole.WriteLine(s.Contains(\"\\n\"));"),
        &["True"]
    );
}

#[test]
fn escape_sequence_newline_produces_newline_character() {
    assert_eq!(
        run_csharp(r#"string s="a\nb"; Console.WriteLine(s.Length);"#),
        &["3"]
    );
}

#[test]
fn escape_sequence_tab_is_single_character() {
    assert_eq!(
        run_csharp(r#"string s="a\tb"; Console.WriteLine(s.Length);"#),
        &["3"]
    );
}

#[test]
fn unicode_escape_produces_correct_character() {
    assert_eq!(
        run_csharp(r#"char c='\u0041'; Console.WriteLine(c);"#),
        &["A"]
    );
}

#[test]
fn raw_string_literal_contains_embedded_quotes_without_escaping() {
    assert_eq!(
        run_csharp(
            r####"string s="""She said "hello" to him.""";
Console.WriteLine(s.Contains("\"hello\""));"####
        ),
        &["True"]
    );
}
