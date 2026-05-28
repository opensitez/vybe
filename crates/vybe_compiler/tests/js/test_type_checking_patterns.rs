/// Type checking, instanceof, typeof, constructor comparison patterns

use super::helpers::run_js;

#[test]
fn typeof_all_primitives_and_object() {
    assert_eq!(run_js(r#"
console.log(typeof undefined);
console.log(typeof null);
console.log(typeof true);
console.log(typeof 42);
console.log(typeof "str");
console.log(typeof Symbol());
console.log(typeof 1n);
console.log(typeof function(){});
console.log(typeof {});
console.log(typeof []);
"#), vec!["undefined", "object", "boolean", "number", "string", "symbol", "bigint", "function", "object", "object"]);
}

#[test]
fn instanceof_class_hierarchy() {
    assert_eq!(run_js(r#"
class A {}
class B extends A {}
class C extends B {}
const obj = new C();
console.log(obj instanceof C);
console.log(obj instanceof B);
console.log(obj instanceof A);
console.log(obj instanceof Object);
"#), vec!["true", "true", "true", "true"]);
}

#[test]
fn constructor_comparison() {
    assert_eq!(run_js(r#"
const arr = [];
const obj = {};
const fn = function() {};
console.log(arr.constructor === Array);
console.log(obj.constructor === Object);
console.log(fn.constructor === Function);
"#), vec!["true", "true", "true"]);
}

#[test]
fn array_check_patterns() {
    assert_eq!(run_js(r#"
const arr = [1, 2, 3];
console.log(Array.isArray(arr));
console.log(Array.isArray({}));
console.log(Array.isArray("string"));
console.log(arr instanceof Array);
"#), vec!["true", "false", "false", "true"]);
}

#[test]
fn duck_typing_check() {
    assert_eq!(run_js(r#"
function isIterable(val) {
    return val != null && typeof val[Symbol.iterator] === "function";
}
console.log(isIterable([1, 2, 3]));
console.log(isIterable("string"));
console.log(isIterable(new Map()));
console.log(isIterable(42));
console.log(isIterable(null));
"#), vec!["true", "true", "true", "false", "false"]);
}

#[test]
fn type_check_with_object_prototype_tostring() {
    assert_eq!(run_js(r#"
function typeOf(val) {
    return Object.prototype.toString.call(val).slice(8, -1);
}
console.log(typeOf(null));
console.log(typeOf(undefined));
console.log(typeOf(42));
console.log(typeOf("str"));
console.log(typeOf([]));
console.log(typeOf({}));
console.log(typeOf(/re/));
"#), vec!["Null", "Undefined", "Number", "String", "Array", "Object", "RegExp"]);
}

#[test]
fn null_undefined_check() {
    assert_eq!(run_js(r#"
function isNullOrUndefined(val) { return val == null; }
console.log(isNullOrUndefined(null));
console.log(isNullOrUndefined(undefined));
console.log(isNullOrUndefined(0));
console.log(isNullOrUndefined(""));
console.log(isNullOrUndefined(false));
"#), vec!["true", "true", "false", "false", "false"]);
}

#[test]
fn plain_object_check() {
    assert_eq!(run_js(r#"
function isPlainObject(val) {
    if (typeof val !== "object" || val === null) return false;
    const proto = Object.getPrototypeOf(val);
    return proto === Object.prototype || proto === null;
}
console.log(isPlainObject({}));
console.log(isPlainObject(Object.create(null)));
console.log(isPlainObject([]));
console.log(isPlainObject(new Date()));
console.log(isPlainObject(42));
"#), vec!["true", "true", "false", "false", "false"]);
}

#[test]
fn number_type_checks() {
    assert_eq!(run_js(r#"
const checks = {
    int: n => Number.isInteger(n),
    finite: n => Number.isFinite(n),
    safe: n => Number.isSafeInteger(n),
    nan: n => Number.isNaN(n),
};
console.log(checks.int(5));
console.log(checks.int(5.5));
console.log(checks.finite(Infinity));
console.log(checks.safe(2**53));
console.log(checks.nan(NaN));
"#), vec!["true", "false", "false", "false", "true"]);
}
