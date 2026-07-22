use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Abstract Equality (`==`, `!=`) Type Coercion Algorithm
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_abstract_equality_null_and_undefined() {
    let src = r#"
console.log(`${null == undefined}:${null == null}:${undefined == undefined}:${null == 0}:${undefined == false}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true:false:false"]); // null == undefined is true, but neither equals 0 or false!
}

#[test]
fn test_js_abstract_equality_string_and_number() {
    let src = r#"
console.log(`${"42" == 42}:${"" == 0}:${"   " == 0}:${"0" == 0}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true:true"]);
}

#[test]
fn test_js_abstract_equality_boolean_coercion_to_number() {
    let src = r#"
console.log(`${true == 1}:${false == 0}:${true == "1"}:${false == ""}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true:true"]);
}

#[test]
fn test_js_abstract_equality_object_and_primitive() {
    let src = r#"
console.log(`${[10] == 10}:${["a"] == "a"}:${{} == "[object Object]"}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_abstract_equality_empty_array_coercions() {
    let src = r#"
console.log(`${[] == false}:${[] == 0}:${[] == ""}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_abstract_equality_bigint_and_number() {
    let src = r#"
console.log(`${10n == 10}:${0n == 0}:${10n == 20}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]);
}

#[test]
fn test_js_abstract_equality_bigint_and_string() {
    let src = r#"
console.log(`${10n == "10"}:${0n == ""}:${5n == " 5 "}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_abstract_equality_nan_is_never_equal_to_anything() {
    let src = r#"
console.log(`${NaN == NaN}:${NaN == 0}:${NaN == false}:${NaN == null}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false:false"]);
}

#[test]
fn test_js_abstract_inequality_operator() {
    let src = r#"
console.log(`${"5" != 5}:${"5" != 6}:${null != undefined}`);
"#;
    assert_eq!(run_js(src), vec!["false:true:false"]);
}

#[test]
fn test_js_abstract_equality_symbol_and_primitive() {
    let src = r#"
const s = Symbol("test");
console.log(`${s == s}:${s == "Symbol(test)"}:${s == 0}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:false"]);
}

#[test]
fn test_js_abstract_equality_symbol_and_object_wrapper() {
    let src = r#"
const s = Symbol("test");
const wrapper = Object(s);
console.log(`${s == wrapper}:${wrapper == s}`);
"#;
    assert_eq!(run_js(src), vec!["true:true"]);
}

#[test]
fn test_js_abstract_equality_custom_valueof_coercion() {
    let src = r#"
const obj = { valueOf: () => 100 };
console.log(obj == 100);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_abstract_equality_custom_tostring_coercion() {
    let src = r#"
const obj = { toString: () => "hello" };
console.log(obj == "hello");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_abstract_equality_toprimitive_overrides_valueof_tostring() {
    let src = r#"
const obj = {
    valueOf: () => 10,
    toString: () => "str",
    [Symbol.toPrimitive]: () => 99
};
console.log(obj == 99);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_abstract_equality_two_objects_compare_by_reference() {
    let src = r#"
const obj1 = { a: 1 };
const obj2 = { a: 1 };
const obj3 = obj1;
console.log(`${obj1 == obj2}:${obj1 == obj3}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]);
}

#[test]
fn test_js_abstract_equality_whitespace_string_number_coercions() {
    let src = r#"
console.log(`${"\t\n" == 0}:${"  12  " == 12}`);
"#;
    assert_eq!(run_js(src), vec!["true:true"]);
}

#[test]
fn test_js_abstract_equality_boolean_and_object() {
    let src = r#"
console.log(`${true == [1]}:${false == []}`);
"#;
    assert_eq!(run_js(src), vec!["true:true"]); // true == [1] -> 1 == [1] -> 1 == 1 -> true!
}

#[test]
fn test_js_abstract_equality_array_with_multiple_elements() {
    let src = r#"
console.log(`${[1, 2] == "1,2"}:${[1, 2] == 1}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_abstract_equality_date_object_default_primitive_hint_is_string() {
    let src = r#"
const d = new Date(0);
console.log(d == d.toString());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_abstract_equality_nested_array_coercion() {
    let src = r#"
console.log(`${[[[5]]] == 5}:${[[[]]] == 0}`);
"#;
    assert_eq!(run_js(src), vec!["true:true"]);
}
