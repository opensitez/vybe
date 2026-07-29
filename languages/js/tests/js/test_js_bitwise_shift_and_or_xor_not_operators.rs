use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Bitwise Shift (`<<`, `>>`, `>>>`) & Bitwise Logic (`&`, `|`, `^`, `~`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_bitwise_and_or_xor_not_operators() {
    let src = r#"
const a = 0b1100, b = 0b1010;
console.log(`${a & b}:${a | b}:${a ^ b}:${(~a & 0xF)}`);
"#;
    assert_eq!(run_js(src), vec!["8:14:6:3"]);
}

#[test]
fn test_js_bitwise_signed_left_right_shift() {
    let src = r#"
console.log(`${5 << 2}:${-5 << 2}:${20 >> 2}:${-20 >> 2}`);
"#;
    assert_eq!(run_js(src), vec!["20:-20:5:-5"]);
}

#[test]
fn test_js_bitwise_unsigned_right_shift_zero_fill() {
    let src = r#"
console.log(`${-1 >>> 0}:${-5 >>> 2}`);
"#;
    assert_eq!(run_js(src), vec!["4294967295:1073741822"]);
}

#[test]
fn test_js_bitwise_shift_amount_modulo_32() {
    let src = r#"
console.log(`${1 << 32}:${1 << 33}:${1 << 35}`); // Shift amount is masked with 0x1F (modulo 32)
"#;
    assert_eq!(run_js(src), vec!["1:2:8"]);
}

#[test]
fn test_js_bitwise_not_double_tilde_truncation() {
    let src = r#"
console.log(`${~~13.37}:${~~-42.84}:${~~NaN}:${~~Infinity}`);
"#;
    assert_eq!(run_js(src), vec!["13:-42:0:0"]);
}

#[test]
fn test_js_bitwise_operators_coerce_operands_to_32bit_int() {
    let src = r#"
console.log(("10" | "5") + "|" + ("15" & "7"));
"#;
    assert_eq!(run_js(src), vec!["15|7"]);
}

#[test]
fn test_js_bitwise_boolean_operand_coercion() {
    let src = r#"
console.log((true & true) + "|" + (true | false) + "|" + (false ^ true));
"#;
    assert_eq!(run_js(src), vec!["1|1|1"]);
}

#[test]
fn test_js_bitwise_null_and_undefined_operand_coercion() {
    let src = r#"
console.log((null | 5) + "|" + (undefined & 10) + "|" + (~null));
"#;
    assert_eq!(run_js(src), vec!["5|0|-1"]);
}

#[test]
fn test_js_bitwise_large_number_wrap_32bit() {
    let src = r#"
const maxInt32PlusOne = 2147483648;
console.log((maxInt32PlusOne | 0) + "|" + (maxInt32PlusOne >>> 0));
"#;
    assert_eq!(run_js(src), vec!["-2147483648|2147483648"]);
}

#[test]
fn test_js_bitwise_assignment_operators() {
    let src = r#"
let x = 0b1010;
x &= 0b1100;
console.log(x);
x |= 0b0001;
console.log(x);
x ^= 0b1001;
console.log(x);
"#;
    assert_eq!(run_js(src), vec!["8", "9", "0"]);
}

#[test]
fn test_js_bitwise_shift_assignment_operators() {
    let src = r#"
let a = 1;
a <<= 3;
console.log(a);
a >>= 1;
console.log(a);
a >>>= 1;
console.log(a);
"#;
    assert_eq!(run_js(src), vec!["8", "4", "2"]);
}

#[test]
fn test_js_bitwise_symbol_operand_throws_typeerror() {
    let src = r#"
try {
    const res = Symbol("a") | 1;
} catch (e) {
    console.log("Bitwise Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Bitwise Symbol TypeError"]);
}

#[test]
fn test_js_bitwise_bigint_and_number_mix_throws_typeerror() {
    let src = r#"
try {
    const res = 10n & 5; // Cannot mix BigInt and Number in bitwise operations!
} catch (e) {
    console.log("Bitwise BigInt Number Mix TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Bitwise BigInt Number Mix TypeError"]);
}

#[test]
fn test_js_bitwise_mask_extraction_pattern() {
    let src = r#"
const flags = 0b1010;
const READ_FLAG = 0b0010;
const WRITE_FLAG = 0b0100;
console.log(`${(flags & READ_FLAG) !== 0}:${(flags & WRITE_FLAG) !== 0}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_bitwise_rgb_color_packing_unpacking() {
    let src = r#"
const r = 255, g = 128, b = 64;
const color = (r << 16) | (g << 8) | b;
const outR = (color >> 16) & 0xFF;
const outG = (color >> 8) & 0xFF;
const outB = color & 0xFF;
console.log(`${outR}:${outG}:${outB}`);
"#;
    assert_eq!(run_js(src), vec!["255:128:64"]);
}

#[test]
fn test_js_bitwise_not_indexof_found_check() {
    let src = r#"
const str = "hello";
console.log((~str.indexOf("e") !== 0) + "|" + (~str.indexOf("z") === 0));
"#;
    assert_eq!(run_js(src), vec!["true|true"]); // ~(-1) is 0
}

#[test]
fn test_js_bitwise_object_valueof_coercion() {
    let src = r#"
const obj1 = { valueOf: () => 12 };
const obj2 = { [Symbol.toPrimitive]: () => 5 };
console.log(obj1 & obj2);
"#;
    assert_eq!(run_js(src), vec!["4"]);
}

#[test]
fn test_js_bitwise_unsigned_right_shift_negative_modulo_32() {
    let src = r#"
console.log(-1 >>> -1); // -1 & 31 = 31 -> -1 >>> 31 = 1
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_bitwise_operator_precedence_with_addition() {
    let src = r#"
console.log((1 + 2 & 3) + "|" + (1 + (2 & 3)));
"#;
    assert_eq!(run_js(src), vec!["3|3"]);
}

#[test]
fn test_js_bitwise_float_truncation_in_shifts() {
    let src = r#"
console.log((10.99 << 1) + "|" + (-10.99 >> 1));
"#;
    assert_eq!(run_js(src), vec!["20|-5"]);
}

#[test]
fn test_js_bitwise_bigint_negative_shift_amount_throws_rangeerror() {
    let src = r#"
try {
    const res = 1n << -1n;
} catch (e) {
    console.log(e.name);
}
"#;
    assert_eq!(run_js(src), vec!["RangeError"]);
}

#[test]
fn test_js_bitwise_not_on_negative_zero() {
    let src = r#"
console.log(~(-0));
"#;
    assert_eq!(run_js(src), vec!["-1"]);
}

