use super::helpers::run_prints;

// ── fdiv — IEEE 754 division ──────────────────────────────────

#[test]
fn fdiv_regular_division() {
    assert_eq!(
        run_prints(r#"<?php echo fdiv(10, 3); "#),
        vec!["3.3333333333333"]
    );
}
#[test]
fn fdiv_divide_by_zero_returns_inf() {
    assert_eq!(run_prints(r#"<?php echo fdiv(1, 0); "#), vec!["INF"]);
}
#[test]
fn fdiv_negative_divide_by_zero() {
    assert_eq!(run_prints(r#"<?php echo fdiv(-1, 0); "#), vec!["-INF"]);
}
#[test]
fn fdiv_zero_divide_by_zero_returns_nan() {
    assert_eq!(run_prints(r#"<?php echo fdiv(0, 0); "#), vec!["NAN"]);
}
#[test]
fn fdiv_integer_exact_result() {
    assert_eq!(run_prints(r#"<?php echo fdiv(10, 2); "#), vec!["5"]);
}

// ── INF and NAN constants ─────────────────────────────────────

#[test]
fn is_nan_detects_nan() {
    assert_eq!(
        run_prints(r#"<?php echo is_nan(NAN) ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn is_nan_rejects_number() {
    assert_eq!(
        run_prints(r#"<?php echo is_nan(42.0) ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}
#[test]
fn is_infinite_detects_inf() {
    assert_eq!(
        run_prints(r#"<?php echo is_infinite(INF) ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn is_infinite_rejects_finite() {
    assert_eq!(
        run_prints(r#"<?php echo is_infinite(1000.0) ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}
#[test]
fn is_finite_rejects_inf() {
    assert_eq!(
        run_prints(r#"<?php echo is_finite(INF) ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}
#[test]
fn is_finite_accepts_float() {
    assert_eq!(
        run_prints(r#"<?php echo is_finite(3.14) ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}

// ── PHP integer limits ────────────────────────────────────────

#[test]
fn php_int_max_is_defined() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_INT_MAX; "#),
        vec!["9223372036854775807"]
    );
}
#[test]
fn php_int_min_is_defined() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_INT_MIN; "#),
        vec!["-9223372036854775808"]
    );
}
#[test]
fn php_int_size_is_eight_bytes() {
    assert_eq!(run_prints(r#"<?php echo PHP_INT_SIZE; "#), vec!["8"]);
}

// ── PHP float constants ───────────────────────────────────────

#[test]
fn php_float_epsilon() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_FLOAT_EPSILON > 0 ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn php_float_max() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_FLOAT_MAX > 1e300 ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}

// ── intdiv edge cases ─────────────────────────────────────────

#[test]
fn intdiv_positive() {
    assert_eq!(run_prints(r#"<?php echo intdiv(7, 2); "#), vec!["3"]);
}
#[test]
fn intdiv_negative_truncates_toward_zero() {
    assert_eq!(run_prints(r#"<?php echo intdiv(-7, 2); "#), vec!["-3"]);
}
#[test]
fn intdiv_exact() {
    assert_eq!(run_prints(r#"<?php echo intdiv(10, 5); "#), vec!["2"]);
}
#[test]
fn intdiv_zero_throws() {
    assert_eq!(
        run_prints(
            r#"<?php
try { intdiv(1, 0); } catch (\DivisionByZeroError $e) { echo "div0"; }
"#
        ),
        vec!["div0"]
    );
}

// ── fmod ──────────────────────────────────────────────────────

#[test]
fn fmod_positive() {
    assert_eq!(run_prints(r#"<?php echo fmod(10.5, 3.0); "#), vec!["1.5"]);
}
#[test]
fn fmod_negative_dividend() {
    assert_eq!(run_prints(r#"<?php echo fmod(-10.5, 3.0); "#), vec!["-1.5"]);
}
#[test]
fn fmod_with_integer_args() {
    assert_eq!(run_prints(r#"<?php echo fmod(10, 3); "#), vec!["1"]);
}

// ── Overflow behavior ─────────────────────────────────────────

#[test]
fn integer_overflow_becomes_float() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = PHP_INT_MAX + 1;
echo is_float($r) ? 'float' : 'int';
"#
        ),
        vec!["float"]
    );
}
#[test]
fn float_addition_precision() {
    assert_eq!(
        run_prints(r#"<?php echo (0.1 + 0.2 < 0.31) ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
