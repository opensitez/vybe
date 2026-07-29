use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Number.isInteger`, `Number.isSafeInteger`, `Number.isNaN`, `Number.isFinite` Utilities
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_number_isinteger_basic() {
    let src = r#"
console.log(`${Number.isInteger(10)}:${Number.isInteger(10.5)}:${Number.isInteger("10")}:${Number.isInteger(NaN)}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:false:false"]);
}

#[test]
fn test_js_number_issafeinteger_range_check() {
    let src = r#"
const maxSafe = Number.MAX_SAFE_INTEGER;
console.log(`${Number.isSafeInteger(maxSafe)}:${Number.isSafeInteger(maxSafe + 1)}:${Number.isSafeInteger(-maxSafe)}:${Number.isSafeInteger(-maxSafe - 1)}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:false"]);
}

#[test]
fn test_js_number_isnan_vs_global_isnan() {
    let src = r#"
console.log(`${Number.isNaN(NaN)}:${Number.isNaN("NaN")}:${isNaN("NaN")}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true"]); // Number.isNaN does NOT perform type coercion!
}

#[test]
fn test_js_number_isfinite_vs_global_isfinite() {
    let src = r#"
console.log(`${Number.isFinite(100)}:${Number.isFinite("100")}:${isFinite("100")}:${Number.isFinite(Infinity)}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:false"]); // Number.isFinite does NOT perform type coercion!
}

#[test]
fn test_js_number_constants_max_min_safe_integers() {
    let src = r#"
console.log(`${Number.MAX_SAFE_INTEGER}:${Number.MIN_SAFE_INTEGER}`);
"#;
    assert_eq!(run_js(src), vec!["9007199254740991:-9007199254740991"]);
}

#[test]
fn test_js_number_isinteger_zero_and_negative_zero() {
    let src = r#"
console.log(Number.isInteger(0) + "|" + Number.isInteger(-0));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_number_isinteger_infinity() {
    let src = r#"
console.log(Number.isInteger(Infinity) + "|" + Number.isInteger(-Infinity));
"#;
    assert_eq!(run_js(src), vec!["false|false"]);
}

#[test]
fn test_js_number_isnan_for_non_numeric_types() {
    let src = r#"
console.log(`${Number.isNaN(undefined)}:${Number.isNaN({})}:${Number.isNaN("hello")}:${Number.isNaN(true)}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false:false"]);
}

#[test]
fn test_js_number_isfinite_for_non_numeric_types() {
    let src = r#"
console.log(`${Number.isFinite(null)}:${Number.isFinite(true)}:${Number.isFinite([])}:${Number.isFinite("0")}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false:false"]);
}

#[test]
fn test_js_number_constants_epsilon() {
    let src = r#"
console.log(Number.EPSILON > 0 && (1 + Number.EPSILON) > 1);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_number_constants_max_min_value() {
    let src = r#"
console.log(Number.MAX_VALUE > 0 && Number.MIN_VALUE > 0);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_number_parseint_parsefloat_on_number_object() {
    let src = r#"
console.log(`${Number.parseInt("42px")}:${Number.parseFloat("3.14159")}`);
"#;
    assert_eq!(run_js(src), vec!["42:3.14159"]);
}

#[test]
fn test_js_number_isinteger_floating_point_with_zero_fraction() {
    let src = r#"
console.log(Number.isInteger(5.0));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_number_issafeinteger_floating_point() {
    let src = r#"
console.log(Number.isSafeInteger(5.0) + "|" + Number.isSafeInteger(5.5));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_number_isnan_expression_evaluations() {
    let src = r#"
console.log(`${Number.isNaN(0 / 0)}:${Number.isNaN(Math.sqrt(-1))}:${Number.isNaN(100)}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]);
}

#[test]
fn test_js_number_isfinite_expression_evaluations() {
    let src = r#"
console.log(`${Number.isFinite(1 / 0)}:${Number.isFinite(-1 / 0)}:${Number.isFinite(Math.log(0))}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false"]);
}

#[test]
fn test_js_number_isinteger_symbol_argument() {
    let src = r#"
console.log(Number.isInteger(Symbol("x")));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_number_isnan_symbol_argument() {
    let src = r#"
console.log(Number.isNaN(Symbol("x")));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_number_parseint_radix_arguments() {
    let src = r#"
console.log(`${Number.parseInt("1010", 2)}:${Number.parseInt("FF", 16)}:${Number.parseInt("077", 8)}`);
"#;
    assert_eq!(run_js(src), vec!["10:255:77"]);
}

#[test]
fn test_js_number_parsefloat_leading_and_trailing_garbage() {
    let src = r#"
console.log(`${Number.parseFloat("  123.456xyz")}:${Number.parseFloat("abc123")}`);
"#;
    assert_eq!(run_js(src), vec!["123.456:NaN"]);
}

#[test]
fn test_js_number_issafeinteger_bigint_argument() {
    let src = r#"
console.log(Number.isSafeInteger(100n));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

