use super::helpers::run_prints;

// ── number_format with 1 arg ──────────────────────────────────

#[test]
fn number_format_integer_no_decimals() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567); "#),
        vec!["1,234,567"]
    );
}

#[test]
fn number_format_float_rounds_to_integer() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234.567); "#),
        vec!["1,235"]
    );
}

#[test]
fn number_format_negative_rounds_correctly() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(-9876543); "#),
        vec!["-9,876,543"]
    );
}

// ── number_format with 2 args (decimals) ─────────────────────

#[test]
fn number_format_two_decimal_places() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234.5, 2); "#),
        vec!["1,234.50"]
    );
}

#[test]
fn number_format_zero_decimal_places_explicit() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(9876.99, 0); "#),
        vec!["9,877"]
    );
}

#[test]
fn number_format_four_decimal_places() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(3.14159265, 4); "#),
        vec!["3.1416"]
    );
}

// ── number_format with 3 args (decimal separator) ────────────

#[test]
fn number_format_comma_decimal_point() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234.56, 2, ','); "#),
        vec!["1,234,56"]
    );
}

#[test]
fn number_format_dot_decimal_no_thousands() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234.56, 2, '.', ''); "#),
        vec!["1234.56"]
    );
}

// ── number_format with 4 args (thousands separator) ──────────

#[test]
fn number_format_european_style() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567.89, 2, ',', '.'); "#),
        vec!["1.234.567,89"]
    );
}

#[test]
fn number_format_space_thousands_separator() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1000000, 0, '.', ' '); "#),
        vec!["1 000 000"]
    );
}

#[test]
fn number_format_underscore_thousands_separator() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1000000.5, 1, '.', '_'); "#),
        vec!["1_000_000.5"]
    );
}

#[test]
fn number_format_empty_thousands_separator() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567, 2, '.', ''); "#),
        vec!["1234567.00"]
    );
}

// ── Edge values ───────────────────────────────────────────────

#[test]
fn number_format_zero_value() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(0, 2); "#),
        vec!["0.00"]
    );
}

#[test]
fn number_format_very_small_float() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(0.001, 3); "#),
        vec!["0.001"]
    );
}

#[test]
fn number_format_large_integer() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(PHP_INT_MAX, 0, '.', ','); "#),
        vec!["9,223,372,036,854,775,807"]
    );
}

#[test]
fn number_format_negative_with_decimals() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(-1234.5678, 2, '.', ','); "#),
        vec!["-1,234.57"]
    );
}

#[test]
fn number_format_negative_european() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(-9876.54, 2, ',', '.'); "#),
        vec!["-9.876,54"]
    );
}

// ── Rounding behavior ─────────────────────────────────────────

#[test]
fn number_format_half_up_rounding() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(2.5, 0); "#),
        vec!["3"]
    );
}

#[test]
fn number_format_rounds_down_below_half() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(2.4, 0); "#),
        vec!["2"]
    );
}

#[test]
fn number_format_five_decimal_precision_loss() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1.23456789, 5); "#),
        vec!["1.23457"]
    );
}

// ── number_format with integer input ─────────────────────────

#[test]
fn number_format_integer_with_decimals_pads_zeros() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(100, 3); "#),
        vec!["100.000"]
    );
}

// ── number_format from computed value ────────────────────────

#[test]
fn number_format_computed_float_result() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1 / 3, 4); "#),
        vec!["0.3333"]
    );
}

#[test]
fn number_format_product_result() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(99.99 * 3, 2); "#),
        vec!["299.97"]
    );
}

// ── money-style formatting ────────────────────────────────────

#[test]
fn number_format_currency_usd_style() {
    assert_eq!(
        run_prints(
            r#"<?php
$amount = 1299.9;
echo '$' . number_format($amount, 2, '.', ',');
"#
        ),
        vec!["$1,299.90"]
    );
}

#[test]
fn number_format_currency_eur_style() {
    assert_eq!(
        run_prints(
            r#"<?php
$amount = 9876.5;
echo number_format($amount, 2, ',', '.') . ' €';
"#
        ),
        vec!["9.876,50 €"]
    );
}
