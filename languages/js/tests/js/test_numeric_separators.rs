/// Numeric separators — underscore separator in decimal, hex, octal, binary,
/// BigInt literals, floating point, scientific notation.
use super::helpers::run_js;

// ── decimal separators ────────────────────────────────────────────────────────

#[test]
fn decimal_separator_basic() {
    assert_eq!(
        run_js(
            r#"
const million = 1_000_000;
console.log(million);
"#
        ),
        vec!["1000000"]
    );
}

#[test]
fn decimal_separator_multiple_groups() {
    assert_eq!(
        run_js(
            r#"
const n = 1_234_567_890;
console.log(n);
"#
        ),
        vec!["1234567890"]
    );
}

#[test]
fn decimal_separator_two_digit_groups() {
    assert_eq!(
        run_js(
            r#"
const n = 1_00;
console.log(n);
"#
        ),
        vec!["100"]
    );
}

// ── float and scientific notation ─────────────────────────────────────────────

#[test]
fn float_separator_in_fractional_part() {
    assert_eq!(
        run_js(
            r#"
const pi = 3.141_592_653;
console.log(pi.toFixed(9));
"#
        ),
        vec!["3.141592653"]
    );
}

#[test]
fn scientific_notation_separator() {
    assert_eq!(
        run_js(
            r#"
const n = 1_000e2;
console.log(n);
"#
        ),
        vec!["100000"]
    );
}

// ── hex separators ────────────────────────────────────────────────────────────

#[test]
fn hex_separator() {
    assert_eq!(
        run_js(
            r#"
const color = 0xFF_FF_FF;
console.log(color);
"#
        ),
        vec!["16777215"]
    );
}

#[test]
fn hex_separator_uint32() {
    assert_eq!(
        run_js(
            r#"
const mask = 0xDEAD_BEEF;
console.log(mask >>> 0);
"#
        ),
        vec!["3735928559"]
    );
}

// ── binary separators ─────────────────────────────────────────────────────────

#[test]
fn binary_separator() {
    assert_eq!(
        run_js(
            r#"
const flags = 0b1010_0001;
console.log(flags);
"#
        ),
        vec!["161"]
    );
}

#[test]
fn binary_separator_byte_groups() {
    assert_eq!(
        run_js(
            r#"
const n = 0b1111_0000_1010_0101;
console.log(n);
"#
        ),
        vec!["61605"]
    );
}

// ── octal separators ──────────────────────────────────────────────────────────

#[test]
fn octal_separator() {
    assert_eq!(
        run_js(
            r#"
const n = 0o777_000;
console.log(n);
"#
        ),
        vec!["261632"]
    );
}

// ── BigInt separators ─────────────────────────────────────────────────────────

#[test]
fn bigint_decimal_separator() {
    assert_eq!(
        run_js(
            r#"
const n = 1_000_000n;
console.log(n === 1000000n);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bigint_hex_separator() {
    assert_eq!(
        run_js(
            r#"
const n = 0xFF_FFn;
console.log(n);
"#
        ),
        vec!["65535n"]
    );
}

// ── separators in expressions ─────────────────────────────────────────────────

#[test]
fn separator_in_computed_expression() {
    assert_eq!(
        run_js(
            r#"
const kb = 1_024;
const mb = 1_024 * 1_024;
console.log(mb / kb);
"#
        ),
        vec!["1024"]
    );
}

#[test]
fn separator_does_not_affect_arithmetic() {
    assert_eq!(
        run_js(
            r#"
console.log(1_000 + 2_000);
console.log(1_0 * 1_0);
"#
        ),
        vec!["3000", "100"]
    );
}

#[test]
fn bigint_binary_separator() {
    assert_eq!(
        run_js(
            r#"
const n = 0b1010_0101n;
console.log(n);
"#
        ),
        vec!["165n"]
    );
}

