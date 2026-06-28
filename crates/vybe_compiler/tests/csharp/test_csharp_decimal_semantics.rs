//! `decimal` is base-10 floating point: arithmetic and comparison preserve
//! scale semantics unlike binary `double`.
use super::helpers::run_csharp;

#[test]
fn decimal_addition_preserves_fractional_sum_without_binary_drift() {
    assert_eq!(
        run_csharp(r#"decimal a = 0.1m; decimal b = 0.2m; Console.WriteLine(a + b);"#),
        &["0.3"]
    );
}

#[test]
fn decimal_subtraction_yields_exact_difference_for_currency_style_values() {
    assert_eq!(
        run_csharp(
            r#"decimal price = 19.99m; decimal discount = 4.50m; Console.WriteLine(price - discount);"#
        ),
        &["15.49"]
    );
}

#[test]
fn decimal_multiplication_scales_both_operands() {
    assert_eq!(
        run_csharp(r#"decimal rate = 1.5m; decimal hours = 2m; Console.WriteLine(rate * hours);"#),
        &["3.0"]
    );
}

#[test]
fn decimal_division_truncates_toward_zero_for_integer_result() {
    assert_eq!(
        run_csharp(r#"decimal total = 10m; decimal parts = 4m; Console.WriteLine(total / parts);"#),
        &["2.5"]
    );
}

#[test]
fn decimal_unary_minus_negates_value() {
    assert_eq!(
        run_csharp(r#"decimal balance = 12.5m; Console.WriteLine(-balance);"#),
        &["-12.5"]
    );
}

#[test]
fn decimal_equality_compares_numeric_value_not_reference() {
    assert_eq!(
        run_csharp(
            r#"
decimal left = 1.0m;
decimal right = 1.00m;
Console.WriteLine(left == right);
"#
        ),
        &["True"]
    );
}

#[test]
fn decimal_comparison_orders_values_before_string_conversion() {
    assert_eq!(
        run_csharp(
            r#"
decimal low = 1.2m;
decimal high = 1.3m;
Console.WriteLine(low < high);
Console.WriteLine(high > low);
"#
        ),
        &["True", "True"]
    );
}

#[test]
fn decimal_modulo_returns_remainder_for_non_integer_division() {
    assert_eq!(run_csharp(r#"Console.WriteLine(10.5m % 3m);"#), &["1.5"]);
}

#[test]
fn decimal_increment_mutates_storage_in_place() {
    assert_eq!(
        run_csharp(
            r#"
decimal tally = 2.5m;
tally++;
Console.WriteLine(tally);
"#
        ),
        &["3.5"]
    );
}

#[test]
fn decimal_cast_from_int_promotes_to_integral_decimal() {
    assert_eq!(
        run_csharp(r#"decimal value = (decimal)7; Console.WriteLine(value);"#),
        &["7"]
    );
}

#[test]
fn decimal_parse_reads_literal_text_without_exponent() {
    assert_eq!(
        run_csharp(r#"decimal value = decimal.Parse("42.5"); Console.WriteLine(value);"#),
        &["42.5"]
    );
}

#[test]
fn decimal_to_string_preserves_trailing_zero_from_format() {
    assert_eq!(
        run_csharp(r#"decimal value = 3.5m; Console.WriteLine(value.ToString("0.00"));"#),
        &["3.50"]
    );
}

#[test]
fn decimal_mixed_addition_with_int_promotes_int_operand() {
    assert_eq!(
        run_csharp(r#"decimal baseAmount = 2.5m; Console.WriteLine(baseAmount + 2);"#),
        &["4.5"]
    );
}

#[test]
fn decimal_zero_is_additive_identity() {
    assert_eq!(
        run_csharp(r#"decimal value = 9.75m; Console.WriteLine(value + 0m);"#),
        &["9.75"]
    );
}

#[test]
fn decimal_one_is_multiplicative_identity() {
    assert_eq!(
        run_csharp(r#"decimal value = 9.75m; Console.WriteLine(value * 1m);"#),
        &["9.75"]
    );
}
