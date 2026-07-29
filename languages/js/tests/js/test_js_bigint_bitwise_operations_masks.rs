use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: BigInt Bitwise Operations (`&`, `|`, `^`, `~`, `<<`, `>>`) & Masking
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_bigint_bitwise_and_or_xor_not() {
    let src = r#"
const a = 0b1100n, b = 0b1010n;
console.log(`${(a & b).toString()}:${(a | b).toString()}:${(a ^ b).toString()}:${(~a).toString()}`);
"#;
    assert_eq!(run_js(src), vec!["8:14:6:-13"]);
}

#[test]
fn test_js_bigint_left_right_shift() {
    let src = r#"
console.log(`${(1n << 64n).toString()}:${(100n >> 2n).toString()}`);
"#;
    assert_eq!(run_js(src), vec!["18446744073709551616:25"]);
}

#[test]
fn test_js_bigint_unsigned_right_shift_prohibited() {
    let src = r#"
try {
    eval("1n >>> 2n;");
} catch (e) {
    console.log("BigInt Unsigned Right Shift TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["BigInt Unsigned Right Shift TypeError"]);
}

#[test]
fn test_js_bigint_asuintn_masking_utility() {
    let src = r#"
const val = 0xFFFFFFFFFFFFFFFFn;
console.log(BigInt.asUintN(16, val).toString());
"#;
    assert_eq!(run_js(src), vec!["65535"]);
}

#[test]
fn test_js_bigint_asintn_sign_extension_utility() {
    let src = r#"
const val = 0xFFFFn; // 65535
console.log(BigInt.asIntN(16, val).toString() + "|" + BigInt.asIntN(32, val).toString());
"#;
    assert_eq!(run_js(src), vec!["-1|65535"]);
}

#[test]
fn test_js_bigint_bitwise_not_negative_numbers() {
    let src = r#"
console.log((~(-10n)).toString()); // ~(-10) = 9
"#;
    assert_eq!(run_js(src), vec!["9"]);
}

#[test]
fn test_js_bigint_bitwise_shift_negative_shift_amount() {
    let src = r#"
console.log((16n << -2n).toString() + "|" + (16n >> -2n).toString()); // Left shift by negative amount is right shift!
"#;
    assert_eq!(run_js(src), vec!["4|64"]);
}

#[test]
fn test_js_bigint_bitwise_assignment_operators() {
    let src = r#"
let x = 0b1111n;
x &= 0b1010n;
console.log(x.toString());
x |= 0b0100n;
console.log(x.toString());
x ^= 0b1111n;
console.log(x.toString());
"#;
    assert_eq!(run_js(src), vec!["10", "14", "1"]);
}

#[test]
fn test_js_bigint_64bit_masking_and_extraction() {
    let src = r#"
const packed = (0x12345678n << 32n) | 0x9ABCDEF0n;
const high = (packed >> 32n) & 0xFFFFFFFFn;
const low = packed & 0xFFFFFFFFn;
console.log(high.toString(16) + "|" + low.toString(16));
"#;
    assert_eq!(run_js(src), vec!["12345678|9abcdef0"]);
}

#[test]
fn test_js_bigint_bitwise_and_with_zero() {
    let src = r#"
console.log((1234567890123456789n & 0n).toString());
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_bigint_bitwise_or_with_zero() {
    let src = r#"
console.log((1234567890123456789n | 0n).toString());
"#;
    assert_eq!(run_js(src), vec!["1234567890123456789"]);
}

#[test]
fn test_js_bigint_asuintn_zero_bits_returns_zero() {
    let src = r#"
console.log(BigInt.asUintN(0, 100n).toString());
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_bigint_asintn_zero_bits_returns_zero() {
    let src = r#"
console.log(BigInt.asIntN(0, 100n).toString());
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_bigint_bitwise_unary_plus_throws_typeerror() {
    let src = r#"
try {
    eval("+10n;");
} catch (e) {
    console.log("Unary Plus BigInt TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Unary Plus BigInt TypeError"]);
}

#[test]
fn test_js_bigint_asuintn_negative_bits_throws_rangeerror() {
    let src = r#"
try {
    BigInt.asUintN(-1, 10n);
} catch (e) {
    console.log("asUintN Negative Bits RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["asUintN Negative Bits RangeError"]);
}

#[test]
fn test_js_bigint_asintn_negative_bits_throws_rangeerror() {
    let src = r#"
try {
    BigInt.asIntN(-1, 10n);
} catch (e) {
    console.log("asIntN Negative Bits RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["asIntN Negative Bits RangeError"]);
}

#[test]
fn test_js_bigint_bitwise_shift_large_amount() {
    let src = r#"
const bigShift = 1n << 1000n;
console.log((bigShift >> 1000n).toString());
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_bigint_bitwise_not_identity() {
    let src = r#"
const val = 987654321n;
console.log((~~val).toString());
"#;
    assert_eq!(run_js(src), vec!["987654321"]);
}

#[test]
fn test_js_bigint_asuintn_non_bigint_target_throws_typeerror() {
    let src = r#"
try {
    BigInt.asUintN(8, 255); // Regular number 255 throws TypeError
} catch (e) {
    console.log("asUintN Non-BigInt Target TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["asUintN Non-BigInt Target TypeError"]);
}

#[test]
fn test_js_bigint_asintn_non_bigint_target_throws_typeerror() {
    let src = r#"
try {
    BigInt.asIntN(8, 255);
} catch (e) {
    console.log("asIntN Non-BigInt Target TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["asIntN Non-BigInt Target TypeError"]);
}

#[test]
fn test_js_bigint_bitwise_shift_assignment_operators() {
    let src = r#"
let x = 1n;
x <<= 4n;
console.log(x.toString());
x >>= 2n;
console.log(x.toString());
"#;
    assert_eq!(run_js(src), vec!["16", "4"]);
}

