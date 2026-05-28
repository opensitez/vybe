/// Math methods — statistics, random (seeded via manual pattern), abs, sign, trunc

use super::helpers::run_js;

#[test]
fn math_abs_various() {
    assert_eq!(run_js(r#"
console.log(Math.abs(-5));
console.log(Math.abs(5));
console.log(Math.abs(-Infinity));
console.log(Math.abs(0));
"#), vec!["5", "5", "Infinity", "0"]);
}

#[test]
fn math_sign_values() {
    assert_eq!(run_js(r#"
console.log(Math.sign(-5));
console.log(Math.sign(0));
console.log(Math.sign(5));
console.log(Math.sign(-0));
console.log(Math.sign(NaN));
"#), vec!["-1", "0", "1", "0", "NaN"]);
}

#[test]
fn math_trunc_vs_floor() {
    assert_eq!(run_js(r#"
console.log(Math.trunc(3.7));
console.log(Math.trunc(-3.7));
console.log(Math.floor(-3.7)); // rounds toward -Infinity
"#), vec!["3", "-3", "-4"]);
}

#[test]
fn math_round_half_up() {
    assert_eq!(run_js(r#"
console.log(Math.round(3.5));
console.log(Math.round(-3.5));
console.log(Math.round(3.4));
"#), vec!["4", "-3", "3"]);
}

#[test]
fn math_clamp_pattern() {
    assert_eq!(run_js(r#"
function clamp(val, min, max) {
    return Math.min(Math.max(val, min), max);
}
console.log(clamp(5, 0, 10));
console.log(clamp(-5, 0, 10));
console.log(clamp(15, 0, 10));
"#), vec!["5", "0", "10"]);
}

#[test]
fn math_hypot_multiple_args() {
    assert_eq!(run_js(r#"
console.log(Math.hypot(3, 4));      // 5
console.log(Math.hypot(5, 12));     // 13
console.log(Math.hypot(1, 1, 1).toFixed(4)); // sqrt(3)
"#), vec!["5", "13", "1.7321"]);
}

#[test]
fn math_log_natural() {
    assert_eq!(run_js(r#"
console.log(Math.log(Math.E).toFixed(10));
console.log(Math.log(1));
console.log(Math.log(0));
"#), vec!["1.0000000000", "0", "-Infinity"]);
}

#[test]
fn math_pow_vs_exponentiation() {
    assert_eq!(run_js(r#"
console.log(Math.pow(2, 10));
console.log(2 ** 10);
console.log(Math.pow(2, 0.5).toFixed(4));
"#), vec!["1024", "1024", "1.4142"]);
}

#[test]
fn math_random_in_range() {
    assert_eq!(run_js(r#"
// Can't assert exact value, but range and type
const r = Math.random();
console.log(typeof r);
console.log(r >= 0 && r < 1);
"#), vec!["number", "true"]);
}

#[test]
fn math_cbrt() {
    assert_eq!(run_js(r#"
console.log(Math.cbrt(27));
console.log(Math.cbrt(-8));
console.log(Math.cbrt(0));
"#), vec!["3", "-2", "0"]);
}

#[test]
fn math_clz32_count_leading_zeros() {
    assert_eq!(run_js(r#"
console.log(Math.clz32(1));    // 31 leading zeros
console.log(Math.clz32(2));    // 30 leading zeros
console.log(Math.clz32(0));    // 32
"#), vec!["31", "30", "32"]);
}

#[test]
fn math_imul_int32_multiply() {
    assert_eq!(run_js(r#"
console.log(Math.imul(3, 4));
console.log(Math.imul(0xffffffff, 5)); // int32 overflow
"#), vec!["12", "-5"]);
}
