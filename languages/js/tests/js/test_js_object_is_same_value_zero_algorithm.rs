use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object.is, SameValue & SameValueZero Algorithms
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_is_nan_equals_nan() {
    let src = r#"
console.log(Object.is(NaN, NaN));
console.log(NaN === NaN);
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_positive_and_negative_zero() {
    let src = r#"
console.log(Object.is(+0, -0));
console.log(+0 === -0);
console.log(Object.is(0, -0));
"#;
    assert_eq!(run_js(src), vec!["false", "true", "false"]);
}

#[test]
fn test_js_object_is_same_references() {
    let src = r#"
const obj1 = { a: 1 };
const obj2 = { a: 1 };
console.log(Object.is(obj1, obj1));
console.log(Object.is(obj1, obj2));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_strings() {
    let src = r#"
console.log(Object.is("hello", "hello"));
console.log(Object.is("hello", "world"));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_booleans() {
    let src = r#"
console.log(Object.is(true, true));
console.log(Object.is(true, false));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_null_and_undefined() {
    let src = r#"
console.log(Object.is(null, null));
console.log(Object.is(undefined, undefined));
console.log(Object.is(null, undefined));
"#;
    assert_eq!(run_js(src), vec!["true", "true", "false"]);
}

#[test]
fn test_js_object_is_symbols() {
    let src = r#"
const s1 = Symbol("id");
const s2 = Symbol("id");
console.log(Object.is(s1, s1));
console.log(Object.is(s1, s2));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_bigint_and_number() {
    let src = r#"
console.log(Object.is(10n, 10n));
console.log(Object.is(10n, 10));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_no_implicit_type_coercion() {
    let src = r#"
console.log(Object.is("0", 0));
console.log(Object.is(false, 0));
console.log(Object.is(null, undefined));
"#;
    assert_eq!(run_js(src), vec!["false", "false", "false"]);
}

#[test]
fn test_js_same_value_zero_array_includes_behavior() {
    let src = r#"
const arr = [NaN, -0];
// Array.prototype.includes uses SameValueZero: NaN equals NaN, and -0 equals +0!
console.log(arr.includes(NaN));
console.log(arr.includes(0));
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}

#[test]
fn test_js_same_value_zero_map_key_lookup() {
    let src = r#"
const map = new Map();
map.set(NaN, "NaN_Value");
map.set(-0, "Zero_Value");

console.log(map.get(NaN));
console.log(map.get(+0));
"#;
    assert_eq!(run_js(src), vec!["NaN_Value", "Zero_Value"]);
}

#[test]
fn test_js_same_value_zero_set_uniqueness() {
    let src = r#"
const set = new Set();
set.add(NaN);
set.add(NaN);
set.add(+0);
set.add(-0);

console.log(set.size);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_object_is_infinities() {
    let src = r#"
console.log(Object.is(Infinity, Infinity));
console.log(Object.is(-Infinity, -Infinity));
console.log(Object.is(Infinity, -Infinity));
"#;
    assert_eq!(run_js(src), vec!["true", "true", "false"]);
}

#[test]
fn test_js_object_is_functions_reference_identity() {
    let src = r#"
function fnA() {}
function fnB() {}
console.log(Object.is(fnA, fnA));
console.log(Object.is(fnA, fnB));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_dates() {
    let src = r#"
const d1 = new Date(2025, 1, 1);
const d2 = new Date(2025, 1, 1);
console.log(Object.is(d1, d1));
console.log(Object.is(d1, d2));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_is_regexes() {
    let src = r#"
const r1 = /abc/g;
const r2 = /abc/g;
console.log(Object.is(r1, r1));
console.log(Object.is(r1, r2));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_same_value_zero_custom_polyfill_simulation() {
    let src = r#"
function sameValueZero(x, y) {
    if (x === y) {
        return true; // Handles +0 and -0 returning true
    }
    return Number.isNaN(x) && Number.isNaN(y);
}
console.log(sameValueZero(NaN, NaN) + "|" + sameValueZero(+0, -0) + "|" + sameValueZero(1, 2));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_object_is_empty_arguments() {
    let src = r#"
console.log(Object.is());
"#;
    assert_eq!(run_js(src), vec!["true"]); // undefined === undefined
}

#[test]
fn test_js_object_is_single_argument() {
    let src = r#"
console.log(Object.is(10));
"#;
    assert_eq!(run_js(src), vec!["false"]); // 10 === undefined -> false
}

#[test]
fn test_js_object_is_typed_array_views() {
    let src = r#"
const buf = new ArrayBuffer(16);
const v1 = new Int32Array(buf);
const v2 = new Int32Array(buf);
console.log(Object.is(v1, v1));
console.log(Object.is(v1, v2));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}
