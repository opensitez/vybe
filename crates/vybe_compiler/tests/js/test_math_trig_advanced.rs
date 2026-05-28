/// Math object — trigonometry, logarithms, hyperbolic functions, sign,
/// trunc, cbrt, clz32, fround, hypot, log2/log10, imul, expm1, log1p.

use super::helpers::run_js;

// ── trigonometry ──────────────────────────────────────────────────────────────

#[test]
fn math_sin_cos_tan() {
    assert_eq!(run_js(r#"
console.log(Math.sin(0));
console.log(Math.cos(0));
console.log(Math.tan(0));
console.log(Math.sin(Math.PI / 2).toFixed(10));
"#), vec!["0", "1", "0", "1.0000000000"]);
}

#[test]
fn math_asin_acos_atan() {
    assert_eq!(run_js(r#"
console.log(Math.asin(1).toFixed(5));
console.log(Math.acos(1).toFixed(5));
console.log(Math.atan(1).toFixed(5));
"#), vec!["1.57080", "0.00000", "0.78540"]);
}

#[test]
fn math_atan2() {
    assert_eq!(run_js(r#"
console.log(Math.atan2(1, 1).toFixed(5));
console.log(Math.atan2(0, 1).toFixed(5));
console.log(Math.atan2(1, 0).toFixed(5));
"#), vec!["0.78540", "0.00000", "1.57080"]);
}

// ── hyperbolic ────────────────────────────────────────────────────────────────

#[test]
fn math_sinh_cosh_tanh() {
    assert_eq!(run_js(r#"
console.log(Math.sinh(0));
console.log(Math.cosh(0));
console.log(Math.tanh(0));
console.log(Math.tanh(Infinity));
console.log(Math.tanh(-Infinity));
"#), vec!["0", "1", "0", "1", "-1"]);
}

#[test]
fn math_asinh_acosh_atanh() {
    assert_eq!(run_js(r#"
console.log(Math.asinh(0));
console.log(Math.acosh(1));
console.log(Math.atanh(0));
"#), vec!["0", "0", "0"]);
}

// ── logarithms ────────────────────────────────────────────────────────────────

#[test]
fn math_log_log2_log10() {
    assert_eq!(run_js(r#"
console.log(Math.log(1));
console.log(Math.log(Math.E).toFixed(10));
console.log(Math.log2(8));
console.log(Math.log10(1000));
"#), vec!["0", "1.0000000000", "3", "3"]);
}

#[test]
fn math_log1p_expm1() {
    assert_eq!(run_js(r#"
console.log(Math.log1p(0));
console.log(Math.expm1(0));
console.log(Math.log1p(Math.E - 1).toFixed(10));
"#), vec!["0", "0", "1.0000000000"]);
}

// ── rounding variants ─────────────────────────────────────────────────────────

#[test]
fn math_trunc() {
    assert_eq!(run_js(r#"
console.log(Math.trunc(4.9));
console.log(Math.trunc(-4.9));
console.log(Math.trunc(0.1));
"#), vec!["4", "-4", "0"]);
}

#[test]
fn math_sign() {
    assert_eq!(run_js(r#"
console.log(Math.sign(5));
console.log(Math.sign(-3));
console.log(Math.sign(0));
console.log(Math.sign(-0));
"#), vec!["1", "-1", "0", "0"]);
}

#[test]
fn math_cbrt() {
    assert_eq!(run_js(r#"
console.log(Math.cbrt(27));
console.log(Math.cbrt(8));
console.log(Math.cbrt(-8));
"#), vec!["3", "2", "-2"]);
}

// ── hypot ─────────────────────────────────────────────────────────────────────

#[test]
fn math_hypot() {
    assert_eq!(run_js(r#"
console.log(Math.hypot(3, 4));
console.log(Math.hypot(5, 12));
console.log(Math.hypot(0));
"#), vec!["5", "13", "0"]);
}

#[test]
fn math_hypot_multiple_args() {
    assert_eq!(run_js(r#"
// 3D distance from origin
const d = Math.hypot(1, 2, 2);
console.log(d);
"#), vec!["3"]);
}

// ── clz32 and imul ────────────────────────────────────────────────────────────

#[test]
fn math_clz32() {
    assert_eq!(run_js(r#"
console.log(Math.clz32(1));      // 31 leading zeros
console.log(Math.clz32(1000));   // fewer leading zeros
console.log(Math.clz32(0));      // 32 (all zeros)
"#), vec!["31", "22", "32"]);
}

#[test]
fn math_imul() {
    assert_eq!(run_js(r#"
console.log(Math.imul(3, 4));
console.log(Math.imul(-1, 8));
console.log(Math.imul(0xffffffff, 5));
"#), vec!["12", "-8", "-5"]);
}

// ── fround ────────────────────────────────────────────────────────────────────

#[test]
fn math_fround_converts_to_float32() {
    assert_eq!(run_js(r#"
// 1.337 can't be exactly represented in float32
const x = Math.fround(1.337);
console.log(x !== 1.337);       // precision differs
console.log(Math.fround(0));
console.log(Math.fround(1));
"#), vec!["true", "0", "1"]);
}

// ── constants ────────────────────────────────────────────────────────────────

#[test]
fn math_constants() {
    assert_eq!(run_js(r#"
console.log(typeof Math.PI);
console.log(typeof Math.E);
console.log(typeof Math.LN2);
console.log(typeof Math.LN10);
console.log(typeof Math.LOG2E);
console.log(typeof Math.LOG10E);
console.log(typeof Math.SQRT2);
console.log(typeof Math.SQRT1_2);
"#), vec!["number", "number", "number", "number", "number", "number", "number", "number"]);
}

#[test]
fn math_pi_precision() {
    assert_eq!(run_js(r#"
console.log(Math.PI.toFixed(5));
"#), vec!["3.14159"]);
}

// ── max/min edge cases ────────────────────────────────────────────────────────

#[test]
fn math_max_min_with_infinity() {
    assert_eq!(run_js(r#"
console.log(Math.max());
console.log(Math.min());
console.log(Math.max(Infinity, -Infinity));
console.log(Math.min(Infinity, -Infinity));
"#), vec!["-Infinity", "Infinity", "Infinity", "-Infinity"]);
}

#[test]
fn math_max_min_with_nan() {
    assert_eq!(run_js(r#"
console.log(isNaN(Math.max(1, NaN, 2)));
console.log(isNaN(Math.min(1, NaN)));
"#), vec!["true", "true"]);
}
