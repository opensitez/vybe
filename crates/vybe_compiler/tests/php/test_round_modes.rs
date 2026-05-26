use super::helpers::run_prints;

// ── round() basic ─────────────────────────────────────────────

#[test] fn round_half_up_positive() {
    assert_eq!(run_prints(r#"<?php echo round(2.5); "#), vec!["3"]);
}
#[test] fn round_half_down_negative() {
    assert_eq!(run_prints(r#"<?php echo round(-2.5); "#), vec!["-3"]);
}
#[test] fn round_below_half_down() {
    assert_eq!(run_prints(r#"<?php echo round(2.4); "#), vec!["2"]);
}
#[test] fn round_above_half_up() {
    assert_eq!(run_prints(r#"<?php echo round(2.6); "#), vec!["3"]);
}

// ── round() with precision ────────────────────────────────────

#[test] fn round_two_decimal_places() {
    assert_eq!(run_prints(r#"<?php echo round(3.14159, 2); "#), vec!["3.14"]);
}
#[test] fn round_negative_precision_tens() {
    assert_eq!(run_prints(r#"<?php echo round(1234.5, -2); "#), vec!["1200"]);
}
#[test] fn round_negative_precision_hundreds() {
    assert_eq!(run_prints(r#"<?php echo round(1550, -3); "#), vec!["2000"]);
}
#[test] fn round_zero_precision_explicit() {
    assert_eq!(run_prints(r#"<?php echo round(9.99, 0); "#), vec!["10"]);
}

// ── PHP_ROUND_HALF_UP ─────────────────────────────────────────

#[test] fn round_half_up_mode() {
    assert_eq!(run_prints(r#"<?php echo round(2.5, 0, PHP_ROUND_HALF_UP); "#), vec!["3"]);
}
#[test] fn round_half_up_negative() {
    assert_eq!(run_prints(r#"<?php echo round(-2.5, 0, PHP_ROUND_HALF_UP); "#), vec!["-3"]);
}

// ── PHP_ROUND_HALF_DOWN ───────────────────────────────────────

#[test] fn round_half_down_mode() {
    assert_eq!(run_prints(r#"<?php echo round(2.5, 0, PHP_ROUND_HALF_DOWN); "#), vec!["2"]);
}
#[test] fn round_half_down_negative_mode() {
    assert_eq!(run_prints(r#"<?php echo round(-2.5, 0, PHP_ROUND_HALF_DOWN); "#), vec!["-2"]);
}

// ── PHP_ROUND_HALF_EVEN (banker's rounding) ───────────────────

#[test] fn round_half_even_rounds_to_nearest_even() {
    assert_eq!(run_prints(r#"<?php echo round(0.5, 0, PHP_ROUND_HALF_EVEN); "#), vec!["0"]);
}
#[test] fn round_half_even_three_point_five() {
    assert_eq!(run_prints(r#"<?php echo round(3.5, 0, PHP_ROUND_HALF_EVEN); "#), vec!["4"]);
}
#[test] fn round_half_even_two_point_five() {
    assert_eq!(run_prints(r#"<?php echo round(2.5, 0, PHP_ROUND_HALF_EVEN); "#), vec!["2"]);
}

// ── PHP_ROUND_HALF_ODD ────────────────────────────────────────

#[test] fn round_half_odd_rounds_to_nearest_odd() {
    assert_eq!(run_prints(r#"<?php echo round(0.5, 0, PHP_ROUND_HALF_ODD); "#), vec!["1"]);
}
#[test] fn round_half_odd_two_point_five() {
    assert_eq!(run_prints(r#"<?php echo round(2.5, 0, PHP_ROUND_HALF_ODD); "#), vec!["3"]);
}

// ── ceil and floor ────────────────────────────────────────────

#[test] fn ceil_positive_fractional() {
    assert_eq!(run_prints(r#"<?php echo ceil(4.1); "#), vec!["5"]);
}
#[test] fn ceil_negative_fractional() {
    assert_eq!(run_prints(r#"<?php echo ceil(-4.9); "#), vec!["-4"]);
}
#[test] fn floor_positive_fractional() {
    assert_eq!(run_prints(r#"<?php echo floor(4.9); "#), vec!["4"]);
}
#[test] fn floor_negative_fractional() {
    assert_eq!(run_prints(r#"<?php echo floor(-4.1); "#), vec!["-5"]);
}
#[test] fn ceil_integer_unchanged() {
    assert_eq!(run_prints(r#"<?php echo ceil(5.0); "#), vec!["5"]);
}
#[test] fn floor_integer_unchanged() {
    assert_eq!(run_prints(r#"<?php echo floor(5.0); "#), vec!["5"]);
}

// ── Round with strings ────────────────────────────────────────

#[test] fn round_numeric_string() {
    assert_eq!(run_prints(r#"<?php echo round("3.7"); "#), vec!["4"]);
}
#[test] fn round_zero() {
    assert_eq!(run_prints(r#"<?php echo round(0.0); "#), vec!["0"]);
}
