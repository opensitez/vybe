//! `char` classification, conversion, and arithmetic.
use super::helpers::run_csharp;

#[test]
fn char_is_digit_true_for_ascii_digit() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(char.IsDigit('7'));"#),
        &["True"]
    );
}

#[test]
fn char_is_letter_true_for_alphabetic_char() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(char.IsLetter('a'));"#),
        &["True"]
    );
}

#[test]
fn char_is_upper_distinguishes_case() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(char.IsUpper('A')); Console.WriteLine(char.IsUpper('a'));"#
        ),
        &["True", "False"]
    );
}

#[test]
fn char_to_upper_converts_lowercase() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(char.ToUpper('b'));"#),
        &["B"]
    );
}

#[test]
fn char_to_lower_converts_uppercase() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(char.ToLower('Z'));"#),
        &["z"]
    );
}

#[test]
fn char_is_whitespace_true_for_space_and_tab() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(char.IsWhiteSpace(' ')); Console.WriteLine(char.IsWhiteSpace('\t'));"#
        ),
        &["True", "True"]
    );
}

#[test]
fn char_is_punctuation_true_for_dot() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(char.IsPunctuation('.'));"#),
        &["True"]
    );
}

#[test]
fn cast_char_to_int_yields_unicode_code_point() {
    assert_eq!(run_csharp(r#"Console.WriteLine((int)'A');"#), &["65"]);
}

#[test]
fn cast_int_to_char_yields_unicode_character() {
    assert_eq!(run_csharp(r#"Console.WriteLine((char)65);"#), &["A"]);
}

#[test]
fn char_arithmetic_adds_offset_to_produce_next_letter() {
    assert_eq!(
        run_csharp(r#"char c = (char)('A' + 2); Console.WriteLine(c);"#),
        &["C"]
    );
}

#[test]
fn char_comparison_uses_unicode_ordinal_ordering() {
    assert_eq!(run_csharp(r#"Console.WriteLine('a' > 'A');"#), &["True"]);
}

#[test]
fn string_from_char_array_roundtrips_via_tochar_array() {
    assert_eq!(
        run_csharp(r#"string s = new string(new char[]{'h','i'}); Console.WriteLine(s);"#),
        &["hi"]
    );
}
