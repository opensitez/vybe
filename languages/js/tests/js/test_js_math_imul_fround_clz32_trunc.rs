use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Math.imul`, `Math.fround`, `Math.clz32`, `Math.trunc` Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_math_imul_32bit_multiplication() {
    let src = r#"
console.log(Math.imul(2, 4) + "|" + Math.imul(-2, 4));
"#;
    assert_eq!(run_js(src), vec!["8|-8"]);
}

#[test]
fn test_js_math_imul_overflow_wrap() {
    let src = r#"
console.log(Math.imul(0xffffffff, 5) + "|" + Math.imul(0xffffffff, 0xffffffff));
"#;
    assert_eq!(run_js(src), vec!["-5|1"]);
}

#[test]
fn test_js_math_fround_single_precision_float() {
    let src = r#"
console.log(Math.fround(1.5) + "|" + (Math.fround(1.337) !== 1.337));
"#;
    assert_eq!(run_js(src), vec!["1.5|true"]);
}

#[test]
fn test_js_math_clz32_leading_zeros() {
    let src = r#"
console.log(`${Math.clz32(1)}:${Math.clz32(1000)}:${Math.clz32(0)}:${Math.clz32(0xffffffff)}`);
"#;
    assert_eq!(run_js(src), vec!["31:22:32:0"]);
}

#[test]
fn test_js_math_trunc_integer_part() {
    let src = r#"
console.log(`${Math.trunc(13.37)}:${Math.trunc(-42.84)}:${Math.trunc(0.123)}:${Math.trunc(-0.123)}`);
"#;
    assert_eq!(run_js(src), vec!["13:-42:0:-0"]);
}

#[test]
fn test_js_math_imul_coercion_to_32bit_int() {
    let src = r#"
console.log(Math.imul("10", "20") + "|" + Math.imul(true, false));
"#;
    assert_eq!(run_js(src), vec!["200|0"]);
}

#[test]
fn test_js_math_clz32_coercion_to_uint32() {
    let src = r#"
console.log(Math.clz32(-1) + "|" + Math.clz32("0x10"));
"#;
    assert_eq!(run_js(src), vec!["0|27"]); // -1 is 0xFFFFFFFF (0 leading zeros), 0x10 is 16 (27 leading zeros)
}

#[test]
fn test_js_math_fround_special_values() {
    let src = r#"
console.log(`${Math.fround(0)}:${Math.fround(-0)}:${Math.fround(Infinity)}:${Math.fround(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["0:-0:Infinity:NaN"]);
}

#[test]
fn test_js_math_trunc_special_values() {
    let src = r#"
console.log(`${Math.trunc(NaN)}:${Math.trunc(Infinity)}:${Math.trunc(-Infinity)}`);
"#;
    assert_eq!(run_js(src), vec!["NaN:Infinity:-Infinity"]);
}

#[test]
fn test_js_math_imul_nan_and_undefined_coercion() {
    let src = r#"
console.log(Math.imul(NaN, 5) + "|" + Math.imul(undefined, 10));
"#;
    assert_eq!(run_js(src), vec!["0|0"]);
}

#[test]
fn test_js_math_clz32_nan_and_undefined_coercion() {
    let src = r#"
console.log(Math.clz32(NaN) + "|" + Math.clz32(undefined));
"#;
    assert_eq!(run_js(src), vec!["32|32"]);
}

#[test]
fn test_js_math_fround_underflow_to_zero() {
    let src = r#"
console.log(Math.fround(1e-50));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_math_fround_overflow_to_infinity() {
    let src = r#"
console.log(`${Math.fround(1e40)}`);
"#;
    assert_eq!(run_js(src), vec!["Infinity"]);
}

#[test]
fn test_js_math_trunc_vs_floor_negative_numbers() {
    let src = r#"
console.log(`${Math.trunc(-3.7)}:${Math.floor(-3.7)}`);
"#;
    assert_eq!(run_js(src), vec!["-3:-4"]);
}

#[test]
fn test_js_math_trunc_string_conversion() {
    let src = r#"
console.log(Math.trunc(" -100.99 "));
"#;
    assert_eq!(run_js(src), vec!["-100"]);
}

#[test]
fn test_js_math_clz32_bitwise_shift_equivalence() {
    let src = r#"
const n = 12345;
console.log(Math.clz32(n) === Math.clz32(n >>> 0));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_math_imul_large_double_inputs() {
    let src = r#"
console.log(Math.imul(1e12, 1e12));
"#;
    assert_eq!(run_js(src), vec!["-1939898368"]);
}

#[test]
fn test_js_math_fround_max_float32() {
    let src = r#"
const maxF32 = 3.4028234663852886e+38;
console.log(Math.fround(maxF32) === maxF32);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_math_trunc_object_with_toprimitive() {
    let src = r#"
const obj = { [Symbol.toPrimitive]: () => "99.5" };
console.log(Math.trunc(obj));
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_math_clz32_object_with_valueof() {
    let src = r#"
const obj = { valueOf: () => 4 };
console.log(Math.clz32(obj));
"#;
    assert_eq!(run_js(src), vec!["29"]);
}

#[test]
fn test_js_math_imul_symbol_argument_throws_typeerror() {
    let src = r#"
try {
    Math.imul(Symbol("id"), 2);
} catch (e) {
    console.log("imul Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["imul Symbol TypeError"]);
}

#[test]
fn test_js_math_clz32_symbol_operand_throws_typeerror() {
    let src = r#"
const errors = [];
try {
    Math.clz32(Symbol("val"));
} catch (e) {
    errors.push("clz32");
}
console.log(errors.join("|"));
"#;
    assert_eq!(run_js(src), vec!["clz32"]);
}

#[test]
fn test_js_math_trunc_symbol_operand_throws_typeerror() {
    let src = r#"
try {
    Math.trunc(Symbol("val"));
} catch (e) {
    console.log(e.name);
}
"#;
    assert_eq!(run_js(src), vec!["TypeError"]);
}

