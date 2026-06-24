use crate::helpers::run_main;

#[test]
fn bigdecimal_add_two_positive_values() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("10.5"); java.math.BigDecimal b = new java.math.BigDecimal("2.5"); System.out.println(a.add(b).toString());"#,
    );
    assert_eq!(out, vec!["13.0"]);
}

#[test]
fn bigdecimal_add_negative_and_positive() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("-3.2"); java.math.BigDecimal b = new java.math.BigDecimal("5.2"); System.out.println(a.add(b).toString());"#,
    );
    assert_eq!(out, vec!["2.0"]);
}

#[test]
fn bigdecimal_subtract_yields_difference() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("10"); java.math.BigDecimal b = new java.math.BigDecimal("4"); System.out.println(a.subtract(b).toString());"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn bigdecimal_subtract_negative_result() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("3"); java.math.BigDecimal b = new java.math.BigDecimal("7"); System.out.println(a.subtract(b).toString());"#,
    );
    assert_eq!(out, vec!["-4"]);
}

#[test]
fn bigdecimal_multiply_integers() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("6"); java.math.BigDecimal b = new java.math.BigDecimal("7"); System.out.println(a.multiply(b).toString());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn bigdecimal_multiply_fractions() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("0.5"); java.math.BigDecimal b = new java.math.BigDecimal("0.2"); System.out.println(a.multiply(b).toString());"#,
    );
    assert_eq!(out, vec!["0.10"]);
}

#[test]
fn bigdecimal_divide_exact_quotient() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("10"); java.math.BigDecimal b = new java.math.BigDecimal("2"); System.out.println(a.divide(b).toString());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn bigdecimal_divide_with_scale_and_half_up() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("1"); java.math.BigDecimal b = new java.math.BigDecimal("3"); System.out.println(a.divide(b, 2, java.math.RoundingMode.HALF_UP).toString());"#,
    );
    assert_eq!(out, vec!["0.33"]);
}

#[test]
fn bigdecimal_divide_half_up_rounds_up_at_five() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("2.5"); java.math.BigDecimal b = new java.math.BigDecimal("1"); System.out.println(a.divide(b, 0, java.math.RoundingMode.HALF_UP).toString());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn bigdecimal_scale_of_integer_value() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("42"); System.out.println(a.scale());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bigdecimal_scale_of_decimal_literal() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("3.140"); System.out.println(a.scale());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn bigdecimal_compare_to_equal_values_is_zero() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("1.0"); java.math.BigDecimal b = new java.math.BigDecimal("1.00"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bigdecimal_compare_to_smaller_is_negative() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("1"); java.math.BigDecimal b = new java.math.BigDecimal("2"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn bigdecimal_compare_to_larger_is_positive() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("9"); java.math.BigDecimal b = new java.math.BigDecimal("3"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bigdecimal_value_of_long_creates_unscaled_integer() {
    let out = run_main(
        r#"java.math.BigDecimal a = java.math.BigDecimal.valueOf(12345L); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn bigdecimal_value_of_double_preserves_decimal_text() {
    let out = run_main(
        r#"java.math.BigDecimal a = java.math.BigDecimal.valueOf(3.14); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn bigdecimal_strip_trailing_zeros_on_fraction() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("1.2000"); System.out.println(a.stripTrailingZeros().toPlainString());"#,
    );
    assert_eq!(out, vec!["1.2"]);
}

#[test]
fn bigdecimal_strip_trailing_zeros_on_integer_with_scale() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("100.00"); System.out.println(a.stripTrailingZeros().toPlainString());"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn bigdecimal_zero_value_to_string() {
    let out = run_main(
        r#"java.math.BigDecimal a = java.math.BigDecimal.ZERO; System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bigdecimal_one_constant_value() {
    let out = run_main(
        r#"java.math.BigDecimal a = java.math.BigDecimal.ONE; System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bigdecimal_ten_constant_value() {
    let out = run_main(
        r#"java.math.BigDecimal a = java.math.BigDecimal.TEN; System.out.println(a.multiply(a).toString());"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn bigdecimal_negate_flips_sign() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("4.5"); System.out.println(a.negate().toString());"#,
    );
    assert_eq!(out, vec!["-4.5"]);
}

#[test]
fn bigdecimal_abs_of_negative_value() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("-8.25"); System.out.println(a.abs().toString());"#,
    );
    assert_eq!(out, vec!["8.25"]);
}

#[test]
fn bigdecimal_plus_same_as_add_identity() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("5"); System.out.println(a.plus().toString());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn bigdecimal_set_scale_with_half_up_rounding() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("2.556"); System.out.println(a.setScale(2, java.math.RoundingMode.HALF_UP).toString());"#,
    );
    assert_eq!(out, vec!["2.56"]);
}

#[test]
fn bigdecimal_move_point_right_multiplies_by_ten() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("1.23"); System.out.println(a.movePointRight(1).toString());"#,
    );
    assert_eq!(out, vec!["12.3"]);
}

#[test]
fn bigdecimal_move_point_left_divides_by_ten() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("12.3"); System.out.println(a.movePointLeft(1).toString());"#,
    );
    assert_eq!(out, vec!["1.23"]);
}

#[test]
fn bigdecimal_signum_positive_value() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("0.01"); System.out.println(a.signum());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bigdecimal_signum_negative_value() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("-0.01"); System.out.println(a.signum());"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn bigdecimal_signum_zero_value() {
    let out = run_main(
        r#"java.math.BigDecimal a = java.math.BigDecimal.ZERO; System.out.println(a.signum());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bigdecimal_unscaled_value_for_integer() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("99"); System.out.println(a.unscaledValue().toString());"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn bigdecimal_precision_counts_significant_digits() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("123.45"); System.out.println(a.precision());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn bigdecimal_max_picks_larger_operand() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("3"); java.math.BigDecimal b = new java.math.BigDecimal("9"); System.out.println(a.max(b).toString());"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn bigdecimal_min_picks_smaller_operand() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("3"); java.math.BigDecimal b = new java.math.BigDecimal("9"); System.out.println(a.min(b).toString());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn bigdecimal_remainder_after_division() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("10"); java.math.BigDecimal b = new java.math.BigDecimal("3"); System.out.println(a.remainder(b).toString());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bigdecimal_equals_ignores_trailing_zero_scale() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("1.0"); java.math.BigDecimal b = new java.math.BigDecimal("1.00"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bigdecimal_compare_to_detects_negative_zero_difference() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("0"); java.math.BigDecimal b = new java.math.BigDecimal("-0"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bigdecimal_add_zero_is_identity() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("77"); System.out.println(a.add(java.math.BigDecimal.ZERO).toString());"#,
    );
    assert_eq!(out, vec!["77"]);
}

#[test]
fn bigdecimal_multiply_by_zero_yields_zero() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("123.456"); System.out.println(a.multiply(java.math.BigDecimal.ZERO).toString());"#,
    );
    assert_eq!(out, vec!["0.000"]);
}

#[test]
fn bigdecimal_divide_by_one_is_unchanged() {
    let out = run_main(
        r#"java.math.BigDecimal a = new java.math.BigDecimal("88.8"); System.out.println(a.divide(java.math.BigDecimal.ONE).toString());"#,
    );
    assert_eq!(out, vec!["88.8"]);
}
