/// BigInt operations — division truncation, exponentiation, bitwise AND/OR/XOR/NOT,
/// left/right shift, negation, unary plus error, Number↔BigInt conversion,
/// BigInt in JSON, BigInt64Array, BigUint64Array.
use super::helpers::run_js;

// ── arithmetic edge cases ─────────────────────────────────────────────────────

#[test]
fn bigint_division_truncates_toward_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(7n / 2n);
console.log(-7n / 2n);
console.log(10n / 3n);
"#
        ),
        vec!["3", "-3", "3"]
    );
}

#[test]
fn bigint_remainder_sign_follows_dividend() {
    assert_eq!(
        run_js(
            r#"
console.log(10n % 3n);
console.log(-10n % 3n);
console.log(10n % -3n);
"#
        ),
        vec!["1", "-1", "1"]
    );
}

#[test]
fn bigint_exponentiation() {
    assert_eq!(
        run_js(
            r#"
console.log(2n ** 10n);
console.log(10n ** 3n);
console.log(2n ** 0n);
"#
        ),
        vec!["1024", "1000", "1"]
    );
}

#[test]
fn bigint_negation() {
    assert_eq!(
        run_js(
            r#"
console.log(-42n);
console.log(-(100n + 1n));
console.log(-(-5n));
"#
        ),
        vec!["-42", "-101", "5"]
    );
}

// ── bitwise ───────────────────────────────────────────────────────────────────

#[test]
fn bigint_bitwise_and() {
    assert_eq!(
        run_js(
            r#"
console.log(0b1100n & 0b1010n);
console.log(0xFF00n & 0x00FFn);
"#
        ),
        vec!["8", "0"]
    );
}

#[test]
fn bigint_bitwise_or() {
    assert_eq!(
        run_js(
            r#"
console.log(0b1100n | 0b1010n);
console.log(0xFF00n | 0x00FFn);
"#
        ),
        vec!["14", "65535"]
    );
}

#[test]
fn bigint_bitwise_xor() {
    assert_eq!(
        run_js(
            r#"
console.log(0b1100n ^ 0b1010n);
console.log(0n ^ 0xFFn);
"#
        ),
        vec!["6", "255"]
    );
}

#[test]
fn bigint_bitwise_not() {
    assert_eq!(
        run_js(
            r#"
console.log(~0n);
console.log(~1n);
console.log(~(-1n));
"#
        ),
        vec!["-1", "-2", "0"]
    );
}

#[test]
fn bigint_left_shift() {
    assert_eq!(
        run_js(
            r#"
console.log(1n << 10n);
console.log(3n << 4n);
"#
        ),
        vec!["1024", "48"]
    );
}

#[test]
fn bigint_right_shift() {
    assert_eq!(
        run_js(
            r#"
console.log(1024n >> 5n);
console.log(-8n >> 1n);
"#
        ),
        vec!["32", "-4"]
    );
}

// ── ordering operators ────────────────────────────────────────────────────────

#[test]
fn bigint_ordering_operators() {
    assert_eq!(
        run_js(
            r#"
console.log(10n < 20n);
console.log(20n > 10n);
console.log(10n <= 10n);
console.log(10n >= 11n);
"#
        ),
        vec!["true", "true", "true", "false"]
    );
}

#[test]
fn bigint_ordering_with_number() {
    assert_eq!(
        run_js(
            r#"
console.log(10n < 20);
console.log(5 > 3n);
"#
        ),
        vec!["true", "true"]
    );
}

// ── type errors ───────────────────────────────────────────────────────────────

#[test]
fn bigint_unary_plus_throws_type_error() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try { const x = +1n; } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── conversion ────────────────────────────────────────────────────────────────

#[test]
fn number_to_bigint_explicit_conversion() {
    assert_eq!(
        run_js(
            r#"
const n = BigInt(Number.MAX_SAFE_INTEGER);
console.log(n === 9007199254740991n);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bigint_to_number_explicit_conversion() {
    assert_eq!(
        run_js(
            r#"
const n = Number(42n);
console.log(n);
console.log(typeof n);
"#
        ),
        vec!["42", "number"]
    );
}

#[test]
fn bigint_in_json_stringify_throws() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try { JSON.stringify(1n); } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── BigInt64Array / BigUint64Array ────────────────────────────────────────────

#[test]
fn bigint64array_stores_bigint_values() {
    assert_eq!(
        run_js(
            r#"
const arr = new BigInt64Array(3);
arr[0] = 100n;
arr[1] = -200n;
arr[2] = 9007199254740993n;
console.log(arr[0]);
console.log(arr[1]);
console.log(arr[2]);
"#
        ),
        vec!["100", "-200", "9007199254740993"]
    );
}

#[test]
fn biguint64array_stores_large_positive() {
    assert_eq!(
        run_js(
            r#"
const arr = new BigUint64Array(1);
arr[0] = 18446744073709551615n;
console.log(arr[0]);
"#
        ),
        vec!["18446744073709551615"]
    );
}

#[test]
fn bigint64array_wraps_on_overflow() {
    assert_eq!(
        run_js(
            r#"
const arr = new BigInt64Array(1);
const max = 9223372036854775807n;
arr[0] = max + 1n;
console.log(arr[0]);
"#
        ),
        vec!["-9223372036854775808"]
    );
}

// ── BigInt.asIntN / BigInt.asUintN ────────────────────────────────────────────

#[test]
fn bigint_as_intn_clamps_signed() {
    assert_eq!(
        run_js(
            r#"
console.log(BigInt.asIntN(8, 255n));
console.log(BigInt.asIntN(8, 128n));
"#
        ),
        vec!["-1", "-128"]
    );
}

#[test]
fn bigint_as_uintn_clamps_unsigned() {
    assert_eq!(
        run_js(
            r#"
console.log(BigInt.asUintN(8, 256n));
console.log(BigInt.asUintN(8, 255n));
"#
        ),
        vec!["0", "255"]
    );
}
