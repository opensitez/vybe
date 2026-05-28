/// delete operator edge cases — own properties, inherited, non-configurable, array elements

use super::helpers::run_js;

#[test]
fn delete_own_property_removes_it() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2 };
delete obj.a;
console.log("a" in obj);
console.log("b" in obj);
"#), vec!["false", "true"]);
}

#[test]
fn delete_returns_true_on_success() {
    assert_eq!(run_js(r#"
const obj = { x: 1 };
console.log(delete obj.x);
"#), vec!["true"]);
}

#[test]
fn delete_non_existent_returns_true() {
    assert_eq!(run_js(r#"
const obj = {};
console.log(delete obj.nope);
"#), vec!["true"]);
}

#[test]
fn delete_inherited_has_no_effect() {
    assert_eq!(run_js(r#"
const proto = { x: 1 };
const obj = Object.create(proto);
delete obj.x; // x is on proto, not own
console.log("x" in obj); // still accessible via prototype
console.log(obj.hasOwnProperty("x"));
"#), vec!["true", "false"]);
}

#[test]
fn delete_array_element_creates_hole() {
    assert_eq!(run_js(r#"
const arr = [1, 2, 3];
delete arr[1];
console.log(arr.length); // length unchanged
console.log(arr[1]);     // undefined (hole)
console.log(1 in arr);   // false — deleted
"#), vec!["3", "undefined", "false"]);
}

#[test]
fn delete_non_configurable_returns_false_in_sloppy() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 1, configurable: false });
const result = delete obj.fixed;
console.log(result);
console.log(obj.fixed);
"#), vec!["false", "1"]);
}

#[test]
fn delete_configurable_defined_property() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperty(obj, "removable", { value: 42, configurable: true });
console.log(delete obj.removable);
console.log("removable" in obj);
"#), vec!["true", "false"]);
}

#[test]
fn delete_var_does_not_work() {
    assert_eq!(run_js(r#"
var x = 1;
const result = delete x;
console.log(result);    // false — vars are non-configurable globals
console.log(typeof x);  // still exists
"#), vec!["false", "number"]);
}

#[test]
fn delete_function_param_does_not_work() {
    assert_eq!(run_js(r#"
function f(param) {
    delete param;
    return typeof param;
}
console.log(f("hello"));
"#), vec!["string"]);
}

#[test]
fn delete_computed_property() {
    assert_eq!(run_js(r#"
const obj = { x: 1, y: 2 };
const key = "x";
delete obj[key];
console.log("x" in obj);
console.log("y" in obj);
"#), vec!["false", "true"]);
}
