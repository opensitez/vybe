/// Number and Math edge cases — NaN, Infinity, -0, float precision, safe integers
use super::helpers::run_js;

#[test]
fn nan_not_equal_to_self() {
    assert_eq!(
        run_js(
            r#"
console.log(NaN === NaN);
console.log(NaN !== NaN);
console.log(Number.isNaN(NaN));
console.log(Object.is(NaN, NaN));
"#
        ),
        vec!["false", "true", "true", "true"]
    );
}

#[test]
fn negative_zero_behavior() {
    assert_eq!(
        run_js(
            r#"
const negZero = -0;
console.log(negZero === 0);
console.log(Object.is(negZero, 0));
console.log(Object.is(negZero, -0));
console.log(String(negZero));
console.log(1 / negZero);
"#
        ),
        vec!["true", "false", "true", "0", "-Infinity"]
    );
}

#[test]
fn infinity_arithmetic() {
    assert_eq!(
        run_js(
            r#"
console.log(Infinity + 1);
console.log(Infinity - Infinity);
console.log(Infinity * -1);
console.log(1 / 0);
console.log(-1 / 0);
"#
        ),
        vec!["Infinity", "NaN", "-Infinity", "Infinity", "-Infinity"]
    );
}

#[test]
fn float_precision_epsilon() {
    assert_eq!(
        run_js(
            r#"
const a = 0.1 + 0.2;
console.log(a === 0.3);
console.log(Math.abs(a - 0.3) < Number.EPSILON);
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn max_safe_integer_boundary() {
    assert_eq!(
        run_js(
            r#"
const max = Number.MAX_SAFE_INTEGER;
console.log(max);
console.log(Number.isSafeInteger(max));
console.log(Number.isSafeInteger(max + 1));
console.log(max + 1 === max + 2); // precision loss
"#
        ),
        vec!["9007199254740991", "true", "false", "true"]
    );
}

#[test]
fn number_constants() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.MAX_SAFE_INTEGER === 2**53 - 1);
console.log(Number.MIN_SAFE_INTEGER === -(2**53 - 1));
console.log(Number.POSITIVE_INFINITY === Infinity);
console.log(Number.NEGATIVE_INFINITY === -Infinity);
console.log(Number.isNaN(Number.NaN));
"#
        ),
        vec!["true", "true", "true", "true", "true"]
    );
}

#[test]
fn tofixed_trailing_zeros() {
    assert_eq!(
        run_js(
            r#"
console.log((1.0).toFixed(2));
console.log((1.5).toFixed(0));
console.log((0).toFixed(3));
"#
        ),
        vec!["1.00", "2", "0.000"]
    );
}

#[test]
fn number_is_integer() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.isInteger(1));
console.log(Number.isInteger(1.0));
console.log(Number.isInteger(1.5));
console.log(Number.isInteger(NaN));
console.log(Number.isInteger(Infinity));
"#
        ),
        vec!["true", "true", "false", "false", "false"]
    );
}

#[test]
fn float_exponent_notation() {
    assert_eq!(
        run_js(
            r#"
console.log(1e3);
console.log(1.5e-2);
console.log(2.5e10);
"#
        ),
        vec!["1000", "0.015", "25000000000"]
    );
}

#[test]
fn number_max_min_value() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.MAX_VALUE > 0);
console.log(Number.MIN_VALUE > 0); // smallest positive, not most negative
console.log(Number.MIN_VALUE < 1);
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn bitwise_not_double_tilde() {
    assert_eq!(
        run_js(
            r#"
// ~~x converts to int32 (truncates)
console.log(~~3.9);
console.log(~~-3.9);
console.log(~~"42");
console.log(~~null);
"#
        ),
        vec!["3", "-3", "42", "0"]
    );
}

#[test]
fn test_infinity_division_returns_nan() {
    assert_eq!(
        run_js(
            r#"
console.log(Infinity / Infinity);
"#
        ),
        vec!["NaN"]
    );
}
