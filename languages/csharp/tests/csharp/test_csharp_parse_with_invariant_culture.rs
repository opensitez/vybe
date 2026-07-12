//! Parsing/formatting with `CultureInfo.InvariantCulture` avoids locale commas.
use super::helpers::run_csharp;

#[test]
fn double_parse_invariant_accepts_dot_decimal_separator() {
    assert_eq!(
        run_csharp(
            r#"
double value = double.Parse("3.5", System.Globalization.CultureInfo.InvariantCulture);
Console.WriteLine(value);
"#
        ),
        &["3.5"]
    );
}

#[test]
fn decimal_to_string_invariant_uses_dot_not_comma_for_fractions() {
    assert_eq!(
        run_csharp(
            r#"
decimal value = 2.25m;
Console.WriteLine(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
"#
        ),
        &["2.25"]
    );
}

#[test]
fn int_parse_invariant_ignores_group_separators_in_strict_mode_failure() {
    assert_eq!(
        run_csharp(
            r#"
try {
    int.Parse("1,234", System.Globalization.CultureInfo.InvariantCulture);
    Console.WriteLine("parsed");
} catch (System.FormatException) {
    Console.WriteLine("reject");
}
"#
        ),
        &["reject"]
    );
}
