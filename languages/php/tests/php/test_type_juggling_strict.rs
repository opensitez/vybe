use super::helpers::run_prints;

// ── declare(strict_types=1) enforcement ──────────────────────

#[test]
fn strict_types_rejects_string_for_int() {
    assert_eq!(
        run_prints(
            r#"<?php
declare(strict_types=1);
function add(int $a, int $b): int { return $a + $b; }
try { echo add('1', 2); }
catch (TypeError $e) { echo 'TypeError'; }
"#
        ),
        vec!["TypeError"]
    );
}
#[test]
fn strict_types_accepts_exact_int() {
    assert_eq!(
        run_prints(
            r#"<?php
declare(strict_types=1);
function square(int $n): int { return $n * $n; }
echo square(5);
"#
        ),
        vec!["25"]
    );
}
#[test]
fn strict_types_float_not_coerced_from_int() {
    assert_eq!(
        run_prints(
            r#"<?php
declare(strict_types=1);
function half(float $n): float { return $n / 2; }
echo half(10.0);
"#
        ),
        vec!["5"]
    );
}

// ── Loose comparison (==) surprises ──────────────────────────

#[test]
fn loose_eq_zero_equals_string() {
    assert_eq!(
        run_prints(r#"<?php echo (0 == 'foo') ? 'equal' : 'not'; "#),
        vec!["not"]
    );
}
#[test]
fn loose_eq_zero_equals_numeric_string() {
    assert_eq!(
        run_prints(r#"<?php echo (0 == '0') ? 'equal' : 'not'; "#),
        vec!["equal"]
    );
}
#[test]
fn loose_eq_string_equals_string() {
    assert_eq!(
        run_prints(r#"<?php echo ('1' == '01') ? 'equal' : 'not'; "#),
        vec!["equal"]
    );
}
#[test]
fn loose_eq_null_equals_empty_string() {
    assert_eq!(
        run_prints(r#"<?php echo (null == '') ? 'equal' : 'not'; "#),
        vec!["equal"]
    );
}
#[test]
fn loose_eq_null_equals_false() {
    assert_eq!(
        run_prints(r#"<?php echo (null == false) ? 'equal' : 'not'; "#),
        vec!["equal"]
    );
}
#[test]
fn strict_eq_null_not_equal_false() {
    assert_eq!(
        run_prints(r#"<?php echo (null === false) ? 'equal' : 'not'; "#),
        vec!["not"]
    );
}

// ── Type juggling in arithmetic ───────────────────────────────

#[test]
fn string_plus_int_converts_string() {
    assert_eq!(run_prints(r#"<?php echo '5 apples' + 3; "#), vec!["8"]);
}
#[test]
fn bool_in_arithmetic() {
    assert_eq!(run_prints(r#"<?php echo true + true + false; "#), vec!["2"]);
}
#[test]
fn null_in_arithmetic() {
    assert_eq!(run_prints(r#"<?php echo null + 5; "#), vec!["5"]);
}
#[test]
fn string_concat_with_int() {
    assert_eq!(
        run_prints(r#"<?php echo 'value: ' . 42; "#),
        vec!["value: 42"]
    );
}

// ── is_* type checking functions ─────────────────────────────

#[test]
fn is_numeric_with_string() {
    assert_eq!(
        run_prints(r#"<?php echo is_numeric('3.14') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn is_numeric_with_hex_string() {
    assert_eq!(
        run_prints(r#"<?php echo is_numeric('0x1A') ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}
#[test]
fn is_int_vs_is_float() {
    assert_eq!(
        run_prints(
            r#"<?php echo is_int(42) ? 'int' : 'not'; echo is_float(3.14) ? 'float' : 'not'; "#
        ),
        vec!["intfloat"]
    );
}
#[test]
fn is_string_is_bool_is_null() {
    assert_eq!(
        run_prints(
            r#"<?php
echo is_string('x') ? '1' : '0';
echo is_bool(false) ? '1' : '0';
echo is_null(null) ? '1' : '0';
"#
        ),
        vec!["111"]
    );
}

// ── Type coercion in comparisons ──────────────────────────────

#[test]
fn comparison_string_to_number() {
    assert_eq!(
        run_prints(r#"<?php echo ('10' > '9') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn comparison_numeric_string_as_number() {
    assert_eq!(
        run_prints(r#"<?php echo ('10' > 9) ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn comparison_null_less_than_everything() {
    assert_eq!(
        run_prints(r#"<?php echo (null < -1) ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}

// ── intval / floatval / strval ────────────────────────────────

#[test]
fn intval_base_16() {
    assert_eq!(run_prints(r#"<?php echo intval('0x1A', 16); "#), vec!["26"]);
}
#[test]
fn intval_base_8() {
    assert_eq!(run_prints(r#"<?php echo intval('017', 8); "#), vec!["15"]);
}
#[test]
fn floatval_string() {
    assert_eq!(run_prints(r#"<?php echo floatval('1.5e2'); "#), vec!["150"]);
}
#[test]
fn strval_various() {
    assert_eq!(
        run_prints(r#"<?php echo strval(3.14) . '|' . strval(true) . '|' . strval(null); "#),
        vec!["3.14|1|"]
    );
}
