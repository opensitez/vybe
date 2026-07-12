//! Numeric format specifiers: D, X, F, E, G, N, P, R, custom patterns.
use super::helpers::run_csharp;

#[test]
fn format_d_pads_integer_with_leading_zeros_to_minimum_width() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(42.ToString("D5"));"#),
        &["00042"]
    );
}

#[test]
fn format_x_lower_encodes_integer_as_lowercase_hex() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(255.ToString("x"));"#),
        &["ff"]
    );
}

#[test]
fn format_x_upper_encodes_integer_as_uppercase_hex() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(255.ToString("X"));"#),
        &["FF"]
    );
}

#[test]
fn format_f2_rounds_double_to_two_decimal_places() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine((3.14159).ToString("F2"));"#),
        &["3.14"]
    );
}

#[test]
fn format_e_expresses_double_in_scientific_notation() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine((1000.0).ToString("E2"));"#),
        &["1.00E+003"]
    );
}

#[test]
fn format_g_chooses_shorter_of_fixed_or_scientific_representation() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine((0.00001).ToString("G"));"#),
        &["1E-05"]
    );
}

#[test]
fn format_n_inserts_group_separators_for_thousands() {
    assert_eq!(
        run_csharp(
            r#"
var s = (1234567).ToString("N0",
    System.Globalization.CultureInfo.InvariantCulture);
Console.WriteLine(s);
"#
        ),
        &["1,234,567"]
    );
}

#[test]
fn format_custom_zero_placeholder_pads_fractional_digits() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine((1.5).ToString("0.00"));"#),
        &["1.50"]
    );
}

#[test]
fn format_custom_hash_placeholder_omits_trailing_zeros() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine((1.5).ToString("0.##"));"#),
        &["1.5"]
    );
}
