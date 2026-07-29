use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Unary Operators (`+`, `-`, `~`, `!`, `void`, `typeof`, `delete`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_unary_plus_numeric_coercion() {
    let src = r#"
console.log(`${+"42"}:${+true}:${+false}:${+null}:${+undefined}:${+""}`);
"#;
    assert_eq!(run_js(src), vec!["42:1:0:0:NaN:0"]);
}

#[test]
fn test_js_unary_minus_numeric_negation() {
    let src = r#"
console.log(`${-"5"}:${-(-10)}:${-true}:${-null}:${-undefined}`);
"#;
    assert_eq!(run_js(src), vec!["-5:10:-1:0:NaN"]);
}

#[test]
fn test_js_unary_bitwise_not_tilde() {
    let src = r#"
console.log(`${~0}:${~5}:${~-1}`);
"#;
    assert_eq!(run_js(src), vec!["-1:-6:0"]);
}

#[test]
fn test_js_unary_logical_not_bang() {
    let src = r#"
console.log(`${!true}:${!false}:${!"hello"}:${!""}:${!0}:${!1}:${!null}:${!undefined}`);
"#;
    assert_eq!(
        run_js(src),
        vec!["false:true:false:true:true:false:true:true"]
    );
}

#[test]
fn test_js_unary_double_logical_not_boolean_coercion() {
    let src = r#"
console.log(`${!!"data"}:${!!0}:${!!{}}:${!!null}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:false"]);
}

#[test]
fn test_js_unary_void_operator_always_returns_undefined() {
    let src = r#"
let sideEffect = 0;
const res = void (sideEffect = 100);
console.log((res === undefined) + "|SideEffect=" + sideEffect);
"#;
    assert_eq!(run_js(src), vec!["true|SideEffect=100"]);
}

#[test]
fn test_js_typeof_all_primitive_types() {
    let src = r#"
console.log([
    typeof undefined,
    typeof null,
    typeof true,
    typeof 42,
    typeof 10n,
    typeof "text",
    typeof Symbol("id"),
    typeof (() => {}),
    typeof {}
].join(","));
"#;
    assert_eq!(
        run_js(src),
        vec!["undefined,object,boolean,number,bigint,string,symbol,function,object"]
    );
}

#[test]
fn test_js_typeof_undeclared_variable_does_not_throw() {
    let src = r#"
console.log(typeof undeclaredNonExistentVariable);
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_delete_object_own_property() {
    let src = r#"
const obj = { a: 1, b: 2 };
const res = delete obj.a;
console.log(res + "|hasA=" + ("a" in obj));
"#;
    assert_eq!(run_js(src), vec!["true|hasA=false"]);
}

#[test]
fn test_js_delete_non_configurable_property_returns_false_in_non_strict() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 1, configurable: false });
const res = delete obj.fixed;
console.log(res + "|fixed=" + obj.fixed);
"#;
    assert_eq!(run_js(src), vec!["false|fixed=1"]);
}

#[test]
fn test_js_delete_array_element_creates_hole() {
    let src = r#"
const arr = [10, 20, 30];
const res = delete arr[1];
console.log(res + "|len=" + arr.length + "|hasHole=" + !(1 in arr));
"#;
    assert_eq!(run_js(src), vec!["true|len=3|hasHole=true"]);
}

#[test]
fn test_js_unary_plus_bigint_throws_typeerror() {
    let src = r#"
try {
    eval("+10n");
} catch (e) {
    console.log("Unary Plus BigInt TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Unary Plus BigInt TypeError"]);
}

#[test]
fn test_js_unary_minus_bigint_negation() {
    let src = r#"
console.log((-10n).toString() + "|" + (-(-5n)).toString());
"#;
    assert_eq!(run_js(src), vec!["-10|5"]);
}

#[test]
fn test_js_unary_bitwise_not_bigint() {
    let src = r#"
console.log((~0n).toString() + "|" + (~5n).toString());
"#;
    assert_eq!(run_js(src), vec!["-1|-6"]);
}

#[test]
fn test_js_unary_plus_object_toprimitive_coercion() {
    let src = r#"
const obj = { [Symbol.toPrimitive]: () => "123" };
console.log(+obj);
"#;
    assert_eq!(run_js(src), vec!["123"]);
}

#[test]
fn test_js_unary_minus_object_valueof_coercion() {
    let src = r#"
const obj = { valueOf: () => 50 };
console.log(-obj);
"#;
    assert_eq!(run_js(src), vec!["-50"]);
}

#[test]
fn test_js_unary_plus_date_object_returns_timestamp() {
    let src = r#"
const d = new Date(1000000000000);
console.log(+d);
"#;
    assert_eq!(run_js(src), vec!["1000000000000"]);
}

#[test]
fn test_js_unary_plus_symbol_throws_typeerror() {
    let src = r#"
try {
    console.log(+Symbol("x"));
} catch (e) {
    console.log("Unary Plus Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Unary Plus Symbol TypeError"]);
}

#[test]
fn test_js_delete_unqualified_variable_in_non_strict() {
    let src = r#"
gVar = 100; // Implicit global variable
const res = delete gVar;
console.log(res + "|hasGlob=" + ("gVar" in globalThis));
"#;
    assert_eq!(run_js(src), vec!["true|hasGlob=false"]);
}

#[test]
fn test_js_delete_declared_var_in_non_strict_returns_false() {
    let src = r#"
var declaredVar = 50;
const res = delete declaredVar;
console.log(res + "|declaredVar=" + declaredVar);
"#;
    assert_eq!(run_js(src), vec!["false|declaredVar=50"]);
}

#[test]
fn test_js_unary_operator_precedence_grouping() {
    let src = r#"
console.log((!+[] ) + "|" + (typeof typeof 123)); // typeof (typeof 123) = typeof "number" = "string"
    "#;
    assert_eq!(run_js(src), vec!["true|string"]);
}

#[test]
fn test_js_delete_nonconfigurable_property_throws_in_strict_mode() {
    let src = r#"
"use strict";
const obj = {};
Object.defineProperty(obj, "fixed", { value: 1, configurable: false });
try {
    delete obj.fixed;
} catch (e) {
    console.log("Delete NonConfigurable Strict TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Delete NonConfigurable Strict TypeError"]);
}

#[test]
fn test_js_void_operator_precedence_with_comma() {
    let src = r#"
let x = 0;
const res = void (x += 1, x += 10);
console.log(`${res === undefined}|${x}`);
"#;
    assert_eq!(run_js(src), vec!["true|11"]);
}

#[test]
fn test_js_delete_undeclared_property_in_strict_mode() {
    let src = r#"
"use strict";
console.log(delete globalThis.noSuchProperty);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_typeof_tdz_variable_throws_referenceerror() {
    let src = r#"
try {
    eval("typeof a; let a = 10;");
} catch (e) {
    console.log(e.name);
}
"#;
    assert_eq!(run_js(src), vec!["ReferenceError"]);
}

#[test]
fn test_js_unary_minus_bigint_zero_is_zero() {
    let src = r#"
console.log((-0n === 0n) + "|" + Object.is(-0n, 0n));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

