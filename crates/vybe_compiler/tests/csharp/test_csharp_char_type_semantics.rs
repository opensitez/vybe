//! `char` is a 16-bit UTF-16 code unit: literals, escapes, and comparisons.
use super::helpers::run_csharp;

#[test]
fn char_literal_single_quote_denotes_code_unit() {
    assert_eq!(
        run_csharp(r#"char letter = 'A'; Console.WriteLine(letter);"#),
        &["A"]
    );
}

#[test]
fn char_escape_tab_produces_whitespace_code_unit() {
    assert_eq!(
        run_csharp(r#"char ch = '\t'; Console.WriteLine(ch == '\t');"#),
        &["True"]
    );
}

#[test]
fn char_escape_newline_matches_linefeed_code_unit() {
    assert_eq!(
        run_csharp(r#"char ch = '\n'; Console.WriteLine((int)ch);"#),
        &["10"]
    );
}

#[test]
fn char_unicode_escape_specifies_code_point() {
    assert_eq!(
        run_csharp(r#"char ch = '\u0041'; Console.WriteLine(ch);"#),
        &["A"]
    );
}

#[test]
fn char_equality_compares_code_units_not_reference() {
    assert_eq!(
        run_csharp(
            r#"
char left = 'Z';
char right = 'Z';
Console.WriteLine(left == right);
"#
        ),
        &["True"]
    );
}

#[test]
fn char_comparison_uses_numeric_code_unit_ordering() {
    assert_eq!(run_csharp(r#"Console.WriteLine('A' < 'B');"#), &["True"]);
}

#[test]
fn char_subtraction_yields_difference_of_code_units() {
    assert_eq!(run_csharp(r#"Console.WriteLine('D' - 'A');"#), &["3"]);
}

#[test]
fn string_indexer_returns_char_at_position() {
    assert_eq!(
        run_csharp(r#"string text = "cat"; Console.WriteLine(text[1]);"#),
        &["a"]
    );
}

#[test]
fn char_to_string_produces_length_one_string() {
    assert_eq!(
        run_csharp(r#"char ch = 'x'; Console.WriteLine(ch.ToString().Length);"#),
        &["1"]
    );
}

#[test]
fn char_is_value_type_stored_on_stack_like_int() {
    assert_eq!(
        run_csharp(
            r#"
char left = 'M';
char right = left;
right = 'N';
Console.WriteLine(left);
"#
        ),
        &["M"]
    );
}

#[test]
fn char_array_holds_sequence_of_code_units() {
    assert_eq!(
        run_csharp(
            r#"
char[] letters = { 'a', 'b', 'c' };
Console.WriteLine(letters[2]);
"#
        ),
        &["c"]
    );
}

#[test]
fn new_string_from_char_array_reconstructs_text() {
    assert_eq!(
        run_csharp(
            r#"
char[] data = { 'h', 'i' };
Console.WriteLine(new string(data));
"#
        ),
        &["hi"]
    );
}
