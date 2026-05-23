use super::helpers::run_js;

// ── Number.isSafeInteger ───────────────────────────────────────────────────

#[test]
fn number_is_safe_integer_with_max_safe() {
    assert_eq!(run_js(r#"
console.log(Number.isSafeInteger(Number.MAX_SAFE_INTEGER));
"#), vec!["true"]);
}

#[test]
fn number_is_safe_integer_max_safe_plus_one_is_false() {
    assert_eq!(run_js(r#"
console.log(Number.isSafeInteger(Number.MAX_SAFE_INTEGER + 1));
"#), vec!["false"]);
}

#[test]
fn number_is_safe_integer_float_is_false() {
    assert_eq!(run_js(r#"
console.log(Number.isSafeInteger(3.5));
"#), vec!["false"]);
}

// ── Number constants ───────────────────────────────────────────────────────

#[test]
fn number_epsilon_is_positive_and_less_than_one() {
    assert_eq!(run_js(r#"
console.log(Number.EPSILON > 0);
console.log(Number.EPSILON < 1);
"#), vec!["true", "true"]);
}

#[test]
fn number_max_safe_integer_value() {
    assert_eq!(run_js(r#"
console.log(Number.MAX_SAFE_INTEGER);
"#), vec!["9007199254740991"]);
}

#[test]
fn number_min_safe_integer_value() {
    assert_eq!(run_js(r#"
console.log(Number.MIN_SAFE_INTEGER);
"#), vec!["-9007199254740991"]);
}

// ── Number.isFinite vs global isFinite ────────────────────────────────────

#[test]
fn number_is_finite_vs_global_is_finite_nan_difference() {
    // global isFinite coerces: isFinite(NaN) = false, isFinite("42") = true
    // Number.isFinite does NOT coerce: Number.isFinite("42") = false
    assert_eq!(run_js(r#"
console.log(Number.isFinite(42));
console.log(Number.isFinite("42"));
console.log(isFinite("42"));
"#), vec!["true", "false", "true"]);
}

// ── Number.isNaN vs global isNaN ─────────────────────────────────────────

#[test]
fn number_is_nan_vs_global_is_nan_string_difference() {
    // global isNaN coerces: isNaN("abc") = true
    // Number.isNaN does NOT coerce: Number.isNaN("abc") = false
    assert_eq!(run_js(r#"
console.log(Number.isNaN(NaN));
console.log(Number.isNaN("abc"));
console.log(isNaN("abc"));
"#), vec!["true", "false", "true"]);
}

// ── Number.parseInt / Number.parseFloat ───────────────────────────────────

#[test]
fn number_parse_int_and_float_same_as_global() {
    assert_eq!(run_js(r#"
console.log(Number.parseInt("10") === parseInt("10"));
console.log(Number.parseFloat("3.14") === parseFloat("3.14"));
"#), vec!["true", "true"]);
}

// ── Number.prototype.toFixed ──────────────────────────────────────────────

#[test]
fn number_to_fixed_various_decimals() {
    assert_eq!(run_js(r#"
let n = 1.23456;
console.log(n.toFixed(0));
console.log(n.toFixed(2));
console.log(n.toFixed(4));
"#), vec!["1", "1.23", "1.2346"]);
}

// ── Number.prototype.toExponential ────────────────────────────────────────

#[test]
fn number_to_exponential() {
    assert_eq!(run_js(r#"
console.log((12345).toExponential(2));
console.log((0.00123).toExponential(2));
"#), vec!["1.23e+4", "1.23e-3"]);
}

// ── Number.prototype.toPrecision ─────────────────────────────────────────

#[test]
fn number_to_precision() {
    assert_eq!(run_js(r#"
console.log((123.456).toPrecision(5));
console.log((0.000123).toPrecision(2));
"#), vec!["123.46", "0.00012"]);
}

// ── Number.prototype.toLocaleString ──────────────────────────────────────

#[test]
fn number_to_locale_string_returns_string() {
    assert_eq!(run_js(r#"
let result = (1234).toLocaleString();
console.log(typeof result);
"#), vec!["string"]);
}

// ── BigInt literal ────────────────────────────────────────────────────────

#[test]
fn bigint_literal() {
    assert_eq!(run_js(r#"
console.log(42n);
"#), vec!["42n"]);
}

// ── BigInt arithmetic ─────────────────────────────────────────────────────

#[test]
fn bigint_arithmetic_add_sub_mul() {
    assert_eq!(run_js(r#"
let a = 10n;
let b = 3n;
console.log(a + b);
console.log(a - b);
console.log(a * b);
"#), vec!["13n", "7n", "30n"]);
}

// ── BigInt comparison ─────────────────────────────────────────────────────

#[test]
fn bigint_comparison_with_bigint() {
    assert_eq!(run_js(r#"
console.log(5n > 3n);
console.log(5n < 3n);
console.log(5n === 5n);
"#), vec!["true", "false", "true"]);
}

// ── BigInt typeof ─────────────────────────────────────────────────────────

#[test]
fn bigint_typeof_returns_bigint() {
    assert_eq!(run_js(r#"
console.log(typeof 42n);
"#), vec!["bigint"]);
}

// ── BigInt.asIntN ─────────────────────────────────────────────────────────

#[test]
fn bigint_as_int_n_clamps_to_signed_n_bit() {
    assert_eq!(run_js(r#"
// 128 in signed 8-bit wraps to -128
console.log(BigInt.asIntN(8, 128n));
// 127 fits in signed 8-bit
console.log(BigInt.asIntN(8, 127n));
"#), vec!["-128n", "127n"]);
}

// ── BigInt.asUintN ────────────────────────────────────────────────────────

#[test]
fn bigint_as_uint_n_clamps_to_unsigned_n_bit() {
    assert_eq!(run_js(r#"
// 256 mod 2^8 = 0
console.log(BigInt.asUintN(8, 256n));
// 255 fits exactly in 8-bit unsigned
console.log(BigInt.asUintN(8, 255n));
"#), vec!["0n", "255n"]);
}

// ── BigInt from string ────────────────────────────────────────────────────

#[test]
fn bigint_from_string() {
    assert_eq!(run_js(r#"
let n = BigInt("9007199254740993");
console.log(n);
"#), vec!["9007199254740993n"]);
}

// ── BigInt toString with radix ────────────────────────────────────────────

#[test]
fn bigint_to_string_with_radix() {
    assert_eq!(run_js(r#"
console.log((255n).toString(16));
console.log((10n).toString(2));
"#), vec!["ff", "1010"]);
}

// ── BigInt via String() ───────────────────────────────────────────────────

#[test]
fn bigint_to_string_via_string_constructor() {
    assert_eq!(run_js(r#"
console.log(String(42n));
"#), vec!["42"]);
}

// ── BigInt mixed arithmetic throws ───────────────────────────────────────

#[test]
fn bigint_mixed_arithmetic_with_number_throws() {
    assert_eq!(run_js(r#"
try {
    let result = 1n + 1;
    console.log("no error");
} catch (e) {
    console.log(e instanceof TypeError);
}
"#), vec!["true"]);
}

// ── BigInt loose vs strict comparison with Number ─────────────────────────

#[test]
fn bigint_loose_vs_strict_comparison_with_number() {
    assert_eq!(run_js(r#"
// == coerces: 1n == 1 is true
console.log(1n == 1);
// === does not coerce: 1n === 1 is false (different types)
console.log(1n === 1);
"#), vec!["true", "false"]);
}

// ── Number.prototype.toString with radix ─────────────────────────────────

#[test]
fn number_to_string_with_radix_hex_and_binary() {
    assert_eq!(run_js(r#"
console.log((255).toString(16));
console.log((10).toString(2));
console.log((8).toString(8));
"#), vec!["ff", "1010", "10"]);
}
