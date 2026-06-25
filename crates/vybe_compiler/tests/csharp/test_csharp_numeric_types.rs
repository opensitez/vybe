//! Numeric type ranges, literals, conversions, overflow.
use super::helpers::run_csharp;

#[test]
fn int_max_value_is_two_billion_one_forty_seven_million() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(int.MaxValue);"#),
        &["2147483647"]
    );
}

#[test]
fn int_min_value_is_negative_two_billion_one_forty_eight_million() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(int.MinValue);"#),
        &["-2147483648"]
    );
}

#[test]
fn long_can_hold_value_beyond_int_max() {
    assert_eq!(
        run_csharp(r#"long x = (long)int.MaxValue + 1; Console.WriteLine(x);"#),
        &["2147483648"]
    );
}

#[test]
fn byte_wraps_to_zero_on_unchecked_overflow() {
    assert_eq!(
        run_csharp(r#"unchecked { byte b = 255; b++; Console.WriteLine(b); }"#),
        &["0"]
    );
}

#[test]
fn uint_holds_value_beyond_int_max() {
    assert_eq!(
        run_csharp(r#"uint u = 3000000000u; Console.WriteLine(u > int.MaxValue);"#),
        &["True"]
    );
}

#[test]
fn float_has_lower_precision_than_double() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(sizeof(float)); Console.WriteLine(sizeof(double));"#),
        &["4", "8"]
    );
}

#[test]
fn decimal_preserves_exact_fractional_value() {
    assert_eq!(
        run_csharp(r#"decimal d = 0.1m + 0.2m; Console.WriteLine(d);"#),
        &["0.3"]
    );
}

#[test]
fn hex_integer_literal_parsed_correctly() {
    assert_eq!(
        run_csharp(r#"int n = 0xFF; Console.WriteLine(n);"#),
        &["255"]
    );
}

#[test]
fn binary_integer_literal_parsed_correctly() {
    assert_eq!(
        run_csharp(r#"int n = 0b1010; Console.WriteLine(n);"#),
        &["10"]
    );
}

#[test]
fn digit_separator_underscore_in_numeric_literal() {
    assert_eq!(
        run_csharp(r#"int million = 1_000_000; Console.WriteLine(million);"#),
        &["1000000"]
    );
}

#[test]
fn short_range_min_max_values() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(short.MinValue); Console.WriteLine(short.MaxValue);"#),
        &["-32768", "32767"]
    );
}
