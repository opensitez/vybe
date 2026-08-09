use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Math.hypot`, `Math.cbrt`, `Math.log1p`, `Math.expm1` Functions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_math_hypot_pythagorean_triples() {
    let src = r#"
console.log(Math.hypot(3, 4) + "|" + Math.hypot(5, 12));
"#;
    assert_eq!(run_js(src), vec!["5|13"]);
}

#[test]
fn test_js_math_hypot_multi_argument_distance() {
    let src = r#"
console.log(Math.hypot(1, 2, 2) + "|" + Math.hypot(2, 3, 6));
"#;
    assert_eq!(run_js(src), vec!["3|7"]);
}

#[test]
fn test_js_math_hypot_overflow_underflow_prevention() {
    let src = r#"
console.log((Math.hypot(1e300, 1e300) > 0) + "|" + (Math.hypot(1e-300, 1e-300) > 0));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_math_cbrt_cube_roots() {
    let src = r#"
console.log(`${Math.cbrt(27)}:${Math.cbrt(-64)}:${Math.cbrt(1)}:${Math.cbrt(0)}`);
"#;
    assert_eq!(run_js(src), vec!["3:-4:1:0"]);
}

#[test]
fn test_js_math_log1p_accurate_log_one_plus_x() {
    let src = r#"
console.log(`${Math.log1p(0)}:${Math.log1p(-1)}:${Math.log1p(1e-15) > 0}`);
"#;
    assert_eq!(run_js(src), vec!["0:-Infinity:true"]);
}

#[test]
fn test_js_math_expm1_accurate_e_to_x_minus_one() {
    let src = r#"
console.log(`${Math.expm1(0)}:${Math.expm1(-Infinity)}:${Math.expm1(1e-15) > 0}`);
"#;
    assert_eq!(run_js(src), vec!["0:-1:true"]);
}

#[test]
fn test_js_math_hypot_zero_arguments() {
    let src = r#"
console.log(Math.hypot());
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_math_hypot_with_infinity_returns_infinity() {
    let src = r#"
console.log(Math.hypot(3, 4, Infinity) + "|" + Math.hypot(-Infinity, 5));
"#;
    assert_eq!(run_js(src), vec!["Infinity|Infinity"]);
}

#[test]
fn test_js_math_hypot_with_nan_returns_nan() {
    let src = r#"
console.log(Math.hypot(3, NaN) + "|" + Math.hypot(NaN, Infinity)); // Infinity takes precedence over NaN
"#;
    assert_eq!(run_js(src), vec!["NaN|Infinity"]);
}

#[test]
fn test_js_math_cbrt_special_values() {
    let src = r#"
console.log(`${Math.cbrt(-0)}:${Math.cbrt(Infinity)}:${Math.cbrt(-Infinity)}:${Math.cbrt(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["-0:Infinity:-Infinity:NaN"]);
}

#[test]
fn test_js_math_log1p_out_of_domain_returns_nan() {
    let src = r#"
console.log(Math.log1p(-2) + "|" + Math.log1p(NaN));
"#;
    assert_eq!(run_js(src), vec!["NaN|NaN"]);
}

#[test]
fn test_js_math_expm1_special_values() {
    let src = r#"
console.log(`${Math.expm1(-0)}:${Math.expm1(Infinity)}:${Math.expm1(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["-0:Infinity:NaN"]);
}

#[test]
fn test_js_math_hypot_coercion_of_string_arguments() {
    let src = r#"
console.log(Math.hypot("3", "4"));
"#;
    assert_eq!(run_js(src), vec!["5"]);
}

#[test]
fn test_js_math_cbrt_string_coercion() {
    let src = r#"
console.log(Math.cbrt("-8"));
"#;
    assert_eq!(run_js(src), vec!["-2"]);
}

#[test]
fn test_js_math_log1p_expm1_inverse_identity() {
    let src = r#"
const x = 0.5;
const restored = Math.log1p(Math.expm1(x));
console.log(restored.toFixed(1));
"#;
    assert_eq!(run_js(src), vec!["0.5"]);
}

#[test]
fn test_js_math_hypot_single_argument() {
    let src = r#"
console.log(Math.hypot(9) + "|" + Math.hypot(-9));
"#;
    assert_eq!(run_js(src), vec!["9|9"]);
}

#[test]
fn test_js_math_cbrt_fractional_input() {
    let src = r#"
console.log(Math.cbrt(0.125));
"#;
    assert_eq!(run_js(src), vec!["0.5"]);
}

#[test]
fn test_js_math_log1p_object_coercion() {
    let src = r#"
const obj = { [Symbol.toPrimitive]: () => 0 };
console.log(Math.log1p(obj));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_math_expm1_object_coercion() {
    let src = r#"
const obj = { valueOf: () => 0 };
console.log(Math.expm1(obj));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_math_hypot_symbol_throws_typeerror() {
    let src = r#"
try {
    Math.hypot(3, Symbol("x"));
} catch (e) {
    console.log("hypot Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["hypot Symbol TypeError"]);
}

#[test]
fn test_js_math_cbrt_symbol_throws_typeerror() {
    let src = r#"
try {
    Math.cbrt(Symbol("x"));
} catch (e) {
    console.log(e.name);
}
"#;
    assert_eq!(run_js(src), vec!["TypeError"]);
}
