use crate::helpers::run_main;

#[test]
fn double_nan_constant_prints_nan() {
    let out = run_main("System.out.println(Double.NaN);");
    assert_eq!(out, vec!["NaN"]);
}

#[test]
fn double_positive_infinity_constant_prints_infinity() {
    let out = run_main("System.out.println(Double.POSITIVE_INFINITY);");
    assert_eq!(out, vec!["Infinity"]);
}

#[test]
fn double_negative_infinity_constant_prints_negative_infinity() {
    let out = run_main("System.out.println(Double.NEGATIVE_INFINITY);");
    assert_eq!(out, vec!["-Infinity"]);
}

#[test]
fn double_is_nan_detects_nan_constant() {
    let out = run_main("System.out.println(Double.isNaN(Double.NaN));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_nan_detects_zero_div_zero_expression() {
    let out = run_main("System.out.println(Double.isNaN(0.0 / 0.0));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_nan_rejects_finite_positive_value() {
    let out = run_main("System.out.println(Double.isNaN(3.14));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_nan_rejects_zero() {
    let out = run_main("System.out.println(Double.isNaN(0.0));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_nan_rejects_positive_infinity() {
    let out = run_main("System.out.println(Double.isNaN(Double.POSITIVE_INFINITY));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_nan_rejects_negative_infinity() {
    let out = run_main("System.out.println(Double.isNaN(Double.NEGATIVE_INFINITY));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_infinite_detects_positive_infinity_constant() {
    let out = run_main("System.out.println(Double.isInfinite(Double.POSITIVE_INFINITY));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_infinite_detects_negative_infinity_constant() {
    let out = run_main("System.out.println(Double.isInfinite(Double.NEGATIVE_INFINITY));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_infinite_detects_one_div_zero() {
    let out = run_main("System.out.println(Double.isInfinite(1.0 / 0.0));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_infinite_rejects_finite_value() {
    let out = run_main("System.out.println(Double.isInfinite(42.0));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_infinite_rejects_nan() {
    let out = run_main("System.out.println(Double.isInfinite(Double.NaN));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_compare_smaller_first_returns_negative_one() {
    let out = run_main("System.out.println(Double.compare(1.5, 2.5));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn double_compare_larger_first_returns_positive_one() {
    let out = run_main("System.out.println(Double.compare(9.0, 3.0));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn double_compare_equal_finite_values_returns_zero() {
    let out = run_main("System.out.println(Double.compare(4.0, 4.0));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_compare_negative_numbers_orders_correctly() {
    let out = run_main("System.out.println(Double.compare(-5.0, -2.0));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn double_compare_negative_to_positive_returns_negative_one() {
    let out = run_main("System.out.println(Double.compare(-1.0, 1.0));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn double_compare_two_nans_returns_zero() {
    let out = run_main("System.out.println(Double.compare(Double.NaN, Double.NaN));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_compare_nan_greater_than_finite_returns_positive_one() {
    let out = run_main("System.out.println(Double.compare(Double.NaN, 1.0));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn double_compare_finite_less_than_nan_returns_negative_one() {
    let out = run_main("System.out.println(Double.compare(1.0, Double.NaN));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn double_compare_positive_infinity_greater_than_max_value() {
    let out = run_main("System.out.println(Double.compare(Double.POSITIVE_INFINITY, Double.MAX_VALUE));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn double_compare_negative_infinity_less_than_min_positive() {
    let out = run_main("System.out.println(Double.compare(Double.NEGATIVE_INFINITY, 1.0));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn double_compare_positive_infinity_equals_positive_infinity() {
    let out = run_main(
        "System.out.println(Double.compare(Double.POSITIVE_INFINITY, Double.POSITIVE_INFINITY));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_compare_negative_infinity_equals_negative_infinity() {
    let out = run_main(
        "System.out.println(Double.compare(Double.NEGATIVE_INFINITY, Double.NEGATIVE_INFINITY));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_compare_positive_infinity_greater_than_finite() {
    let out = run_main("System.out.println(Double.compare(Double.POSITIVE_INFINITY, 1000.0));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn double_compare_positive_zero_greater_than_negative_zero() {
    let out = run_main("System.out.println(Double.compare(+0.0, -0.0));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn double_compare_negative_zero_less_than_positive_zero() {
    let out = run_main("System.out.println(Double.compare(-0.0, +0.0));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn double_compare_positive_zero_equals_positive_zero() {
    let out = run_main("System.out.println(Double.compare(+0.0, +0.0));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_nan_not_equal_to_itself_via_equality() {
    let out = run_main("System.out.println(Double.NaN == Double.NaN);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_nan_not_equal_to_finite_value() {
    let out = run_main("System.out.println(Double.NaN == 1.0);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_positive_infinity_greater_than_finite_via_comparison() {
    let out = run_main("System.out.println(Double.POSITIVE_INFINITY > 1.0e308);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_negative_infinity_less_than_finite_via_comparison() {
    let out = run_main("System.out.println(Double.NEGATIVE_INFINITY < -1.0e308);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_nan_after_sqrt_of_negative() {
    let out = run_main("System.out.println(Double.isNaN(Math.sqrt(-1.0)));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_finite_on_regular_fraction() {
    let out = run_main("System.out.println(Double.isFinite(2.5));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_finite_rejects_nan() {
    let out = run_main("System.out.println(Double.isFinite(Double.NaN));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_finite_rejects_positive_infinity() {
    let out = run_main("System.out.println(Double.isFinite(Double.POSITIVE_INFINITY));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_compare_subnormal_values_orders_by_magnitude() {
    let out = run_main("System.out.println(Double.compare(1.0e-300, 2.0e-300));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn double_compare_same_negative_finite_values_returns_zero() {
    let out = run_main("System.out.println(Double.compare(-7.25, -7.25));");
    assert_eq!(out, vec!["0"]);
}
