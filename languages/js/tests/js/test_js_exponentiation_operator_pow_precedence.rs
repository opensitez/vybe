use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Exponentiation Operator (`**`, `**=`) & Math.pow Precedence
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_exponentiation_basic_power() {
    let src = r#"
console.log((2 ** 3) + "|" + (3 ** 2) + "|" + (10 ** 0));
"#;
    assert_eq!(run_js(src), vec!["8|9|1"]);
}

#[test]
fn test_js_exponentiation_right_associativity() {
    let src = r#"
console.log((2 ** 3 ** 2) + "|" + (2 ** (3 ** 2)));
"#;
    assert_eq!(run_js(src), vec!["512|512"]); // Right associative: 2 ** (3 ** 2) = 2 ** 9 = 512
}

#[test]
fn test_js_exponentiation_unary_operator_syntax_error() {
    let src = r#"
try {
    eval("-2 ** 2"); // Unary minus before ** is a SyntaxError without parentheses!
} catch (e) {
    console.log("Unary Minus Exponentiation SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Unary Minus Exponentiation SyntaxError"]);
}

#[test]
fn test_js_exponentiation_with_parenthesized_unary_minus() {
    let src = r#"
console.log(((-2) ** 2) + "|" + (-(2 ** 2)));
"#;
    assert_eq!(run_js(src), vec!["4|-4"]);
}

#[test]
fn test_js_exponentiation_fractional_power_square_root() {
    let src = r#"
console.log((16 ** 0.5) + "|" + (27 ** (1/3)).toFixed(1));
"#;
    assert_eq!(run_js(src), vec!["4|3.0"]);
}

#[test]
fn test_js_exponentiation_negative_exponent() {
    let src = r#"
console.log((2 ** -2) + "|" + (10 ** -1));
"#;
    assert_eq!(run_js(src), vec!["0.25|0.1"]);
}

#[test]
fn test_js_exponentiation_assignment_operator() {
    let src = r#"
let base = 3;
base **= 3;
console.log(base);
"#;
    assert_eq!(run_js(src), vec!["27"]);
}

#[test]
fn test_js_exponentiation_bigint_powers() {
    let src = r#"
console.log((2n ** 100n).toString());
"#;
    assert_eq!(run_js(src), vec!["1267650600228229401496703205376"]);
}

#[test]
fn test_js_exponentiation_bigint_negative_exponent_throws_rangeerror() {
    let src = r#"
try {
    eval("2n ** -1n");
} catch (e) {
    console.log("BigInt Negative Exponent RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["BigInt Negative Exponent RangeError"]);
}

#[test]
fn test_js_exponentiation_nan_and_infinity() {
    let src = r#"
console.log(`${NaN ** 2}:${2 ** NaN}:${Infinity ** 0}:${1 ** Infinity}`);
"#;
    assert_eq!(run_js(src), vec!["NaN:NaN:1:NaN"]);
}

#[test]
fn test_js_exponentiation_coercion_of_string_operands() {
    let src = r#"
console.log(("3" ** "3") + "|" + ("4" ** 0.5));
"#;
    assert_eq!(run_js(src), vec!["27|2"]);
}

#[test]
fn test_js_exponentiation_boolean_operands() {
    let src = r#"
console.log((true ** 5) + "|" + (false ** 0) + "|" + (2 ** true));
"#;
    assert_eq!(run_js(src), vec!["1|1|2"]);
}

#[test]
fn test_js_exponentiation_null_and_undefined_operands() {
    let src = r#"
console.log((null ** 2) + "|" + (undefined ** 2) + "|" + (2 ** null));
"#;
    assert_eq!(run_js(src), vec!["0|NaN|1"]);
}

#[test]
fn test_js_exponentiation_symbol_operand_throws_typeerror() {
    let src = r#"
try {
    Symbol("2") ** 2;
} catch (e) {
    console.log("Exponentiation Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Exponentiation Symbol TypeError"]);
}

#[test]
fn test_js_exponentiation_precedence_relative_to_multiplication() {
    let src = r#"
console.log((2 * 3 ** 2) + "|" + ((2 * 3) ** 2)); // ** has higher precedence than *
"#;
    assert_eq!(run_js(src), vec!["18|36"]);
}

#[test]
fn test_js_exponentiation_precedence_relative_to_addition() {
    let src = r#"
console.log((1 + 2 ** 3) + "|" + ((1 + 2) ** 3));
"#;
    assert_eq!(run_js(src), vec!["9|27"]);
}

#[test]
fn test_js_exponentiation_math_pow_equivalence() {
    let src = r#"
console.log((Math.pow(5, 3) === (5 ** 3)) + "|" + (Math.pow(2, 0.5) === (2 ** 0.5)));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_exponentiation_object_valueof_coercion() {
    let src = r#"
const obj = { valueOf: () => 4 };
console.log(obj ** 2);
"#;
    assert_eq!(run_js(src), vec!["16"]);
}

#[test]
fn test_js_exponentiation_toprimitive_coercion() {
    let src = r#"
const obj = { [Symbol.toPrimitive]: () => 3 };
console.log(2 ** obj);
"#;
    assert_eq!(run_js(src), vec!["8"]);
}

#[test]
fn test_js_exponentiation_mixed_bigint_number_throws_typeerror() {
    let src = r#"
try {
    eval("2n ** 3");
} catch (e) {
    console.log("Mixed BigInt Number Exponentiation TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Mixed BigInt Number Exponentiation TypeError"]
    );
}

#[test]
fn test_js_exponentiation_zero_and_negative_zero_sign_preservation() {
    let src = r#"
console.log(`${0 ** -1}:${(-0) ** -1}:${(-0) ** -2}:${Object.is((-0) ** 3, -0)}:${Object.is((-0) ** 2, 0)}`);
"#;
    assert_eq!(run_js(src), vec!["Infinity:-Infinity:Infinity:true:true"]);
}

#[test]
fn test_js_exponentiation_infinity_and_fractional_base_boundaries() {
    let src = r#"
console.log(`${(-2) ** Infinity}:${(0.5) ** Infinity}:${(-0.5) ** Infinity}:${(-1) ** Infinity}`);
"#;
    assert_eq!(run_js(src), vec!["Infinity:0:0:NaN"]);
}

#[test]
fn test_js_exponentiation_assignment_on_accessor_property() {
    let src = r#"
const obj = {
    _val: 2,
    get val() { return this._val; },
    set val(v) { this._val = v; }
};
obj.val **= 3;
console.log(`${obj.val}:${obj._val}`);
"#;
    assert_eq!(run_js(src), vec!["8:8"]);
}

#[test]
fn test_js_exponentiation_bigint_zero_exponent_and_large_powers() {
    let src = r#"
console.log(`${(0n ** 0n).toString()}:${(5n ** 0n).toString()}:${(2n ** 64n).toString()}`);
"#;
    assert_eq!(run_js(src), vec!["1:1:18446744073709551616"]);
}


