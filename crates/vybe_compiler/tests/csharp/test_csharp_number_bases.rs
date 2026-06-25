//! Numeric literal bases: hex, binary, octal emulation, `Convert` base conversions.
use super::helpers::run_csharp;

#[test]
fn hex_literal_represents_correct_decimal_value() {
    assert_eq!(run_csharp(r#"Console.WriteLine(0xFF);"#), &["255"]);
}

#[test]
fn binary_literal_represents_correct_decimal_value() {
    assert_eq!(run_csharp(r#"Console.WriteLine(0b1010);"#), &["10"]);
}

#[test]
fn underscore_separator_does_not_change_value() {
    assert_eq!(run_csharp(r#"Console.WriteLine(1_000_000);"#), &["1000000"]);
}

#[test]
fn convert_to_string_with_base_16_formats_hex() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Convert.ToString(255,16));"#),
        &["ff"]
    );
}

#[test]
fn convert_to_string_with_base_2_formats_binary() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Convert.ToString(10,2));"#),
        &["1010"]
    );
}

#[test]
fn convert_from_base_16_string_to_int() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Convert.ToInt32("ff",16));"#),
        &["255"]
    );
}

#[test]
fn long_hex_literal_covers_full_64_bit_range() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(0x7FFFFFFFFFFFFFFFL==long.MaxValue);"#),
        &["True"]
    );
}
