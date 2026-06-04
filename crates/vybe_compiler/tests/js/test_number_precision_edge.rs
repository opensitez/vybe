/// Number edge cases — Number.isInteger/isFinite/isNaN vs global,
/// Number.parseFloat/parseInt, toFixed/toPrecision/toExponential,
/// integer limits, signed zero, IEEE 754 quirks.
use super::helpers::run_js;

// ── Number static methods ────────────────────────────────────────────────────

#[test]
fn number_is_integer() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.isInteger(42));
console.log(Number.isInteger(42.0));
console.log(Number.isInteger(42.5));
console.log(Number.isInteger(Infinity));
console.log(Number.isInteger(NaN));
"#
        ),
        vec!["true", "true", "false", "false", "false"]
    );
}

#[test]
fn number_is_finite_vs_global() {
    assert_eq!(
        run_js(
            r#"
// Number.isFinite does NOT coerce
console.log(Number.isFinite(42));
console.log(Number.isFinite(Infinity));
console.log(Number.isFinite("42")); // false — no coercion
console.log(isFinite("42"));        // true — coerces
"#
        ),
        vec!["true", "false", "false", "true"]
    );
}

#[test]
fn number_is_nan_vs_global() {
    assert_eq!(
        run_js(
            r#"
// Number.isNaN does NOT coerce
console.log(Number.isNaN(NaN));
console.log(Number.isNaN("NaN")); // false — no coercion
console.log(isNaN("NaN"));        // true — coerces
console.log(Number.isNaN(undefined));
"#
        ),
        vec!["true", "false", "true", "false"]
    );
}

#[test]
fn number_is_safe_integer() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.isSafeInteger(42));
console.log(Number.isSafeInteger(Number.MAX_SAFE_INTEGER));
console.log(Number.isSafeInteger(Number.MAX_SAFE_INTEGER + 1));
console.log(Number.isSafeInteger(3.5));
"#
        ),
        vec!["true", "true", "false", "false"]
    );
}

// ── Number constants ──────────────────────────────────────────────────────────

#[test]
fn number_max_safe_integer() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.MAX_SAFE_INTEGER);
console.log(Number.MIN_SAFE_INTEGER);
console.log(Number.MAX_SAFE_INTEGER === 2 ** 53 - 1);
"#
        ),
        vec!["9007199254740991", "-9007199254740991", "true"]
    );
}

#[test]
fn number_epsilon() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.EPSILON > 0);
console.log(Number.EPSILON < 0.001);
// Using epsilon for float comparison
const a = 0.1 + 0.2;
const b = 0.3;
console.log(Math.abs(a - b) < Number.EPSILON * 10);
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn number_max_min_value() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.MAX_VALUE > 0);
console.log(Number.MIN_VALUE > 0);
console.log(Number.MIN_VALUE < 0.001);
console.log(Number.POSITIVE_INFINITY === Infinity);
console.log(Number.NEGATIVE_INFINITY === -Infinity);
"#
        ),
        vec!["true", "true", "true", "true", "true"]
    );
}

// ── Number.parseInt/parseFloat ────────────────────────────────────────────────

#[test]
fn number_parseint_radix() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.parseInt("ff", 16));
console.log(Number.parseInt("10", 2));
console.log(Number.parseInt("777", 8));
console.log(Number.parseInt("42abc"));
"#
        ),
        vec!["255", "2", "511", "42"]
    );
}

#[test]
fn number_parsefloat_basics() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.parseFloat("3.14"));
console.log(Number.parseFloat("3.14xyz"));
console.log(isNaN(Number.parseFloat("abc")));
"#
        ),
        vec!["3.14", "3.14", "true"]
    );
}

// ── toFixed / toPrecision / toExponential ─────────────────────────────────────

#[test]
fn to_fixed_rounds_decimal() {
    assert_eq!(
        run_js(
            r#"
console.log((3.14159).toFixed(2));
console.log((1.005).toFixed(2)); // IEEE 754 — might be 1.00 or 1.01
console.log((100).toFixed(2));
"#
        ),
        vec!["3.14", "1.00", "100.00"]
    );
}

#[test]
fn to_precision_significant_digits() {
    assert_eq!(
        run_js(
            r#"
console.log((123.456).toPrecision(5));
console.log((0.000123).toPrecision(2));
console.log((1).toPrecision(4));
"#
        ),
        vec!["123.46", "0.00012", "1.000"]
    );
}

#[test]
fn to_exponential() {
    assert_eq!(
        run_js(
            r#"
console.log((12345).toExponential(2));
console.log((0.00123).toExponential(2));
console.log((1).toExponential(0));
"#
        ),
        vec!["1.23e+4", "1.23e-3", "1e+0"]
    );
}

// ── signed zero ───────────────────────────────────────────────────────────────

#[test]
fn negative_zero_equals_positive_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(-0 === 0);
console.log(-0 === -0);
console.log(Object.is(-0, 0));
console.log(Object.is(-0, -0));
"#
        ),
        vec!["true", "true", "false", "true"]
    );
}

#[test]
fn negative_zero_string_representation() {
    assert_eq!(
        run_js(
            r#"
console.log(String(-0));
console.log((-0).toString());
console.log(JSON.stringify(-0));
"#
        ),
        vec!["0", "0", "0"]
    );
}

#[test]
fn negative_zero_arithmetic() {
    assert_eq!(
        run_js(
            r#"
console.log(1 / -0);
console.log(1 / 0);
console.log(-1 * 0);
"#
        ),
        vec!["-Infinity", "Infinity", "-0"]
    );
}

// ── IEEE 754 quirks ───────────────────────────────────────────────────────────

#[test]
fn floating_point_rounding_error() {
    assert_eq!(
        run_js(
            r#"
console.log(0.1 + 0.2 === 0.3);
console.log(0.1 + 0.2);
"#
        ),
        vec!["false", "0.30000000000000004"]
    );
}

#[test]
fn number_to_string_removes_trailing_zeros() {
    assert_eq!(
        run_js(
            r#"
console.log((1.0).toString());
console.log((1.50).toString());
console.log((1.500000).toString());
"#
        ),
        vec!["1", "1.5", "1.5"]
    );
}

// ── parseInt edge cases ───────────────────────────────────────────────────────

#[test]
fn parseint_stops_at_invalid_character() {
    assert_eq!(
        run_js(
            r#"
console.log(parseInt("10.5"));
console.log(parseInt("0xFF"));
console.log(parseInt(""));
console.log(isNaN(parseInt("")));
"#
        ),
        vec!["10", "255", "NaN", "true"]
    );
}

#[test]
fn parseint_leading_whitespace_trimmed() {
    assert_eq!(
        run_js(
            r#"
console.log(parseInt("  42  "));
"#
        ),
        vec!["42"]
    );
}
